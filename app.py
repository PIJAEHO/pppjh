"""
Render / Railway 등 클라우드 배포용 진입점.
gunicorn app:app 으로 실행됩니다.
"""
import importlib.util, os

spec = importlib.util.spec_from_file_location(
    "server",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "환율서버.py"),
)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
app = mod.app
