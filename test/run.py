"""밴드매니저 시험 실행기 — node 로 test/*.test.js 를 돌린다 (브라우저·서버 없이)
   python test/run.py"""
import glob, os, subprocess, sys

NODE = r"C:\Program Files\nodejs\node.exe"
here = os.path.dirname(os.path.abspath(__file__))
bad = 0
for f in sorted(glob.glob(os.path.join(here, "*.test.js"))):
    r = subprocess.run([NODE, f], cwd=os.path.dirname(here))
    bad += r.returncode != 0
sys.exit(1 if bad else 0)
