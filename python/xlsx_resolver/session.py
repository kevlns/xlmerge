"""审阅服务的持久化身份与完成回执。"""
import json
import os
import uuid
import time
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.parse import urlparse


def save_session(file, data):
    file = Path(file)
    file.parent.mkdir(parents=True, exist_ok=True)
    temp = file.with_name(file.name + '.' + uuid.uuid4().hex + '.tmp')
    temp.write_text(json.dumps(data, ensure_ascii=False), encoding='utf-8')
    os.replace(temp, file)


def inspect_session(manifest, cancel=False):
    file = Path(manifest).resolve().parent / 'session.json'
    original = json.loads(file.read_text(encoding='utf-8'))
    if str(Path(original['manifest']).resolve()) != str(Path(manifest).resolve()):
        raise ValueError('审阅身份与清单不一致')
    if original['state'] in ('completed', 'cancelled', 'failed'):
        return {**original, 'ok': True, 'stopConfirmed': True}
    url = urlparse(original['url'])
    if url.scheme != 'http' or url.hostname != '127.0.0.1' or not url.port or url.path:
        raise ValueError('审阅服务地址无效')
    try:
        live = json.load(urlopen(original['url'] + '/api/session', timeout=3))
        if live.get('sessionId') != original['sessionId'] or live.get('pid') != original['pid']:
            raise ValueError('审阅服务身份变化，拒绝操作')
        if not cancel:
            return {**live, 'ok': True, 'stopConfirmed': False}
        req = Request(original['url'] + '/api/shutdown', data=b'{}', headers={'X-Xlmerge-Session': original['sessionId']}, method='POST')
        with urlopen(req, timeout=3):
            pass
        deadline = time.monotonic() + 5
        while time.monotonic() < deadline:
            settled = json.loads(file.read_text(encoding='utf-8'))
            if settled['state'] in ('completed', 'cancelled', 'failed'):
                return {**settled, 'ok': True, 'stopConfirmed': True}
            time.sleep(.1)
        return {**live, 'ok': True, 'state': 'cancelling', 'stopConfirmed': False}
    except OSError as exc:
        settled = json.loads(file.read_text(encoding='utf-8'))
        if settled['state'] in ('completed', 'cancelled', 'failed'):
            return {**settled, 'ok': True, 'stopConfirmed': True}
        return {**original, 'ok': False, 'state': 'unknown', 'stopConfirmed': False, 'error': str(exc)}
