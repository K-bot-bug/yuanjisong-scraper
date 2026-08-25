"""
代理池 - 自动抓取 + 连通测试 + 轮转 + 失败剔除 + 自动恢复
========================================================
功能：
- 从多个免费源（快代理、free-proxy-list.net）抓取代理列表
- 用 curl_cffi 测试每个代理能否访问目标站点
- 维护工作池/黑名单池，失败剔除，定期恢复
- 每次请求随机选取一个有效代理

使用方式：
    from proxy_pool import ProxyPool
    pool = ProxyPool(target_url='https://www.yuanjisong.com/job/allcity')
    await pool.init()        # 初始化（抓取+测试）
    proxy = pool.get_proxy() # 获取一个可用代理
"""
import asyncio
import random
import re
import sys
import time
from typing import Optional

sys.stdout.reconfigure(encoding='utf-8')

from curl_cffi import requests as cffi_requests
from loguru import logger

# ========== 配置 ==========
PROXY_SOURCES = [
    # 快代理（国内）
    "https://www.kuaiproxy.com/proxy/1/",
    # free-proxy-list.net（国际）
    "https://free-proxy-list.net/",
]

# 每页抓取的代理数量
PROXIES_PER_SOURCE = 50
# 代理最大失败次数后拉黑
BLACKLIST_THRESHOLD = 3
# 拉黑代理恢复检查间隔（秒）
RETRY_BLACKLIST_INTERVAL = 300
# 重新抓取代理间隔（秒）
REFRESH_INTERVAL = 1800  # 30分钟

TARGET_TEST_URL = "https://www.yuanjisong.com/job/allcity"
TEST_TIMEOUT = 8


class Proxy:
    """单个代理记录"""
    __slots__ = ('ip', 'port', 'proto', 'url', 'success_count', 'fail_count',
                 'last_success_time', 'last_fail_time', 'blocked', 'source')

    def __init__(self, ip: str, port: int, proto: str = 'http', source: str = ''):
        self.ip = ip
        self.port = port
        self.proto = proto
        self.url = f"{proto}://{ip}:{port}"
        self.success_count = 0
        self.fail_count = 0
        self.last_success_time = 0.0
        self.last_fail_time = 0.0
        self.blocked = False
        self.source = source

    def __repr__(self):
        return f"{self.url} (ok={self.success_count}, fail={self.fail_count})"

    def to_dict(self):
        return {'url': self.url, 'source': self.source,
                'success': self.success_count, 'fail': self.fail_count}


class ProxyPool:
    """
    异步代理池
    - init() 启动后台刷新任务
    - get_proxy() 获取一个可用代理（无则返回 None）
    - mark_success(proxy) / mark_fail(proxy) 更新代理状态
    - close() 停止后台任务
    """

    def __init__(self, target_url: str = TARGET_TEST_URL):
        self._target = target_url
        self._working: list[Proxy] = []
        self._blacklist: list[Proxy] = []
        self._lock = asyncio.Lock()
        self._running = False
        self._refresh_task: Optional[asyncio.Task] = None
        self._stats = {'fetched': 0, 'tested': 0, 'valid': 0, 'invalid': 0}

    # ============ 代理抓取 ============

    @staticmethod
    def _parse_proxy_table(html: str) -> list[tuple[str, int, str]]:
        """解析 HTML 表格中的代理 IP:Port（通用解析器）"""
        results = []
        rows = re.findall(r'<tr>(.*?)</tr>', html, re.DOTALL)
        for row in rows:
            cols = re.findall(r'<td[^>]*>(.*?)</td>', row, re.DOTALL)
            if len(cols) >= 2:
                ip_m = re.search(r'(\d+\.\d+\.\d+\.\d+)', cols[0])
                port_m = re.search(r'(\d+)', cols[1])
                if ip_m and port_m:
                    ip, port = ip_m.group(1), int(port_m.group(1))
                    if 1 <= port <= 65535:
                        results.append((ip, port, 'http'))
        return results

    async def _fetch_proxies_from_source(self, session: cffi_requests.Session,
                                         source_url: str) -> list[Proxy]:
        """从单个源抓取代理"""
        proxies = []
        try:
            resp = session.get(source_url, timeout=10, allow_redirects=True)
            if resp.status_code != 200:
                logger.debug(f"抓取代理源失败: {source_url} ({resp.status_code})")
                return proxies
            parsed = self._parse_proxy_table(resp.text)
            for ip, port, proto in parsed[:PROXIES_PER_SOURCE]:
                proxies.append(Proxy(ip, port, proto, source_url))
            logger.debug(f"从 {source_url} 抓到 {len(proxies)} 个代理")
        except Exception as e:
            logger.debug(f"抓取代理源异常 {source_url}: {e}")
        return proxies

    async def _fetch_all_proxies(self, session: cffi_requests.Session) -> list[Proxy]:
        """从所有源抓取代理（合并去重）"""
        all_proxies = {}
        for url in PROXY_SOURCES:
            proxies = await self._fetch_proxies_from_source(session, url)
            for p in proxies:
                key = f"{p.ip}:{p.port}"
                if key not in all_proxies:
                    all_proxies[key] = p
        return list(all_proxies.values())

    # ============ 代理测试 ============

    async def _test_proxy(self, session: cffi_requests.AsyncSession, proxy: Proxy) -> bool:
        """测试单个代理是否能访问目标站点"""
        try:
            resp = await session.get(
                self._target,
                proxies={'http': proxy.url, 'https': proxy.url},
                impersonate='chrome131',
                timeout=TEST_TIMEOUT,
                allow_redirects=True
            )
            return resp.status_code == 200 and 'job_card' in resp.text
        except Exception:
            return False

    async def _test_batch(self, proxies: list[Proxy]) -> tuple[list[Proxy], list[Proxy]]:
        """批量测试代理，返回 (working, blacklist) 列表"""
        working, blacklist = [], []
        async with cffi_requests.AsyncSession() as session:
            tasks = [self._test_proxy(session, p) for p in proxies]
            results = await asyncio.gather(*tasks, return_exceptions=True)

            for proxy, ok in zip(proxies, results):
                if ok is True:
                    proxy.success_count += 1
                    proxy.last_success_time = time.time()
                    working.append(proxy)
                elif ok is False:
                    proxy.fail_count += 1
                    proxy.last_fail_time = time.time()
                    if proxy.fail_count >= BLACKLIST_THRESHOLD:
                        proxy.blocked = True
                        blacklist.append(proxy)
        return working, blacklist

    # ============ 公共接口 ============

    async def init(self) -> dict:
        """
        初始化代理池：抓取 → 测试 → 加入工作池
        返回统计信息
        """
        logger.info("开始初始化代理池...")
        start = time.time()

        session = cffi_requests.Session()
        try:
            # 1. 抓取
            raw_proxies = await self._fetch_all_proxies(session)
            self._stats['fetched'] = len(raw_proxies)
            logger.info(f"抓取到 {len(raw_proxies)} 个原始代理")

            if not raw_proxies:
                logger.warning("未抓到任何代理，尝试使用直连")
                return {'fetched': 0, 'working': 0, 'elapsed': time.time() - start}

            # 2. 测试
            working, blacklist = await self._test_batch(raw_proxies)
            self._stats['tested'] = len(raw_proxies)
            self._stats['valid'] = len(working)
            self._stats['invalid'] = len(blacklist)

            # 3. 合并到工作池（保留历史计数）
            async with self._lock:
                existing = {f"{p.ip}:{p.port}": p for p in self._working}
                for p in working:
                    key = f"{p.ip}:{p.port}"
                    if key in existing:
                        existing[key].success_count = max(existing[key].success_count, p.success_count)
                    else:
                        self._working.append(p)
                for p in blacklist:
                    if not p.blocked:
                        p.blocked = True
                    self._blacklist.append(p)
        finally:
            session.close()

        elapsed = time.time() - start
        logger.info(f"代理池初始化完成: {len(self._working)} 个可用 / {len(self._blacklist)} 个黑名单 "
                     f"(抓取{self._stats['fetched']}, 耗时{elapsed:.1f}s)")
        return {'fetched': self._stats['fetched'],
                'working': len(self._working),
                'elapsed': round(elapsed, 1)}

    def get_proxy(self) -> Optional[str]:
        """
        获取一个可用代理 URL（随机选取）
        如果工作池为空，返回 None（调用方用直连）
        """
        if not self._working:
            return None
        # 按成功率加权随机选取
        weights = [max(1, p.success_count) for p in self._working]
        total = sum(weights)
        r = random.uniform(0, total)
        cumulative = 0
        for p in self._working:
            cumulative += max(1, p.success_count)
            if r <= cumulative:
                return p.url
        return self._working[-1].url

    async def mark_success(self, proxy_url: str):
        """标记代理成功"""
        async with self._lock:
            for p in self._working:
                if p.url == proxy_url:
                    p.success_count += 1
                    p.last_success_time = time.time()
                    p.blocked = False
                    break

    async def mark_fail(self, proxy_url: str):
        """标记代理失败，达到阈值加入黑名单"""
        async with self._lock:
            for p in self._working:
                if p.url == proxy_url:
                    p.fail_count += 1
                    p.last_fail_time = time.time()
                    if p.fail_count >= BLACKLIST_THRESHOLD:
                        p.blocked = True
                        self._working.remove(p)
                        self._blacklist.append(p)
                    break
            # 同时尝试从黑名单恢复
            for p in list(self._blacklist):
                if p.blocked and p.fail_count >= BLACKLIST_THRESHOLD:
                    age = time.time() - p.last_fail_time
                    if age > RETRY_BLACKLIST_INTERVAL:
                        p.blocked = False
                        p.fail_count = 0
                        self._blacklist.remove(p)
                        self._working.append(p)

    async def _background_refresh(self):
        """后台定期刷新代理池"""
        while self._running:
            await asyncio.sleep(REFRESH_INTERVAL)
            if not self._running:
                break
            logger.info("后台刷新代理池...")
            await self.init()

    async def start_background(self):
        """启动后台刷新任务"""
        self._running = True
        self._refresh_task = asyncio.create_task(self._background_refresh())
        logger.info(f"代理池后台刷新已启动 (间隔{REFRESH_INTERVAL}s)")

    async def stop(self):
        """停止后台任务"""
        self._running = False
        if self._refresh_task:
            self._refresh_task.cancel()
            try:
                await self._refresh_task
            except asyncio.CancelledError:
                pass
        logger.info("代理池已停止")

    def status(self) -> dict:
        """查看代理池状态"""
        return {
            'working': len(self._working),
            'blacklist': len(self._blacklist),
            'stats': dict(self._stats),
            'top_proxies': [p.to_dict() for p in sorted(self._working,
                                                        key=lambda p: p.success_count, reverse=True)[:5]],
        }


# ========== 模块级单例 ==========
_pool: Optional[ProxyPool] = None


def get_pool() -> ProxyPool:
    """获取全局代理池实例"""
    global _pool
    if _pool is None:
        _pool = ProxyPool()
    return _pool


async def init_pool() -> dict:
    """初始化并启动后台刷新"""
    pool = get_pool()
    result = await pool.init()
    await pool.start_background()
    return result


def get_proxy_url() -> Optional[str]:
    """获取一个可用代理 URL"""
    pool = get_pool()
    return pool.get_proxy()
