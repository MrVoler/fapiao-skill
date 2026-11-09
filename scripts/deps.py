# -*- coding: utf-8 -*-
"""依赖自举：缺什么装什么，不打断用户。

首次运行自动 pip install 缺失的包；装不上才提示手动命令。
pywin32 只在 Windows 上尝试，且属可选——缺了只是不能实测页数 / 导 PDF。
"""
import importlib
import os
import subprocess
import sys

# (模块名, pip 包名, 是否必需)
DEPS = [
    ('docx', 'python-docx', True),
    ('pymupdf', 'pymupdf', True),
    ('PIL', 'pillow', True),
    ('pypdf', 'pypdf', True),
    ('win32com.client', 'pywin32', False),
]

# 先走默认源，失败再用国内镜像兜底
INDEXES = [None, 'https://pypi.tuna.tsinghua.edu.cn/simple']


def _importable(mod):
    """在当前进程里试导入，快。"""
    try:
        importlib.import_module(mod)
        return True
    except Exception:
        return False


def _importable_fresh(mod):
    """起个新进程试导入。安装后用这个复查，避免本进程的残留状态误判。"""
    try:
        r = subprocess.run([sys.executable, '-c', f'import {mod}'],
                           capture_output=True, timeout=120)
        return r.returncode == 0
    except Exception:
        return False


def _missing(fresh=False):
    out = []
    for mod, pkg, required in DEPS:
        if pkg == 'pywin32' and os.name != 'nt':
            continue
        ok = _importable_fresh(mod) if fresh else _importable(mod)
        if not ok:
            out.append((pkg, required))
    return out


def _run_pip(pkgs, index):
    cmd = [sys.executable, '-m', 'pip', 'install',
           '--disable-pip-version-check', '--no-input', '--quiet']
    if index:
        cmd += ['--index-url', index]
    cmd += pkgs
    try:
        return subprocess.run(cmd, capture_output=True, text=True, timeout=900)
    except (OSError, subprocess.TimeoutExpired) as e:
        return subprocess.CompletedProcess(cmd, 1, '', str(e))


def _pip_install(pkgs):
    """装一组包，返回 (是否成功, 最后的错误摘要)。"""
    last = ''
    for index in INDEXES:
        r = _run_pip(pkgs, index)
        if r.returncode == 0:
            return True, ''
        last = (r.stderr or r.stdout or '').strip()
    return False, last


def ensure(quiet=False):
    """确保依赖可用。缺必需包且装不上时直接退出。"""
    todo = _missing()
    if not todo:
        return

    pkgs = [p for p, _ in todo]
    if not quiet:
        print(f'[依赖] 首次使用，自动安装 {len(pkgs)} 个组件：{"、".join(pkgs)}', flush=True)

    ok, err = _pip_install(pkgs)

    # 有些环境 pip 不在 PATH 上，补一次
    if not ok:
        try:
            subprocess.run([sys.executable, '-m', 'ensurepip', '--upgrade'],
                           capture_output=True, timeout=300)
        except Exception:
            pass
        ok, err = _pip_install(pkgs)

    still = _missing(fresh=True)
    if not still:
        if not quiet:
            print('[依赖] 装好了，继续。', flush=True)
        return

    need = [p for p, req in still if req]
    optional = [p for p, req in still if not req]

    if optional and not quiet:
        print(f'[提示] {"、".join(optional)} 没装上，将跳过页数实测与 PDF 导出。', flush=True)

    if need:
        print('[依赖] 自动安装失败，请手动执行一次：', file=sys.stderr)
        print(f'  {sys.executable} -m pip install {" ".join(need)}', file=sys.stderr)
        if err:
            print(f'  原因：{err.splitlines()[-1][:200]}', file=sys.stderr)
        sys.exit(2)
