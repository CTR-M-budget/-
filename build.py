"""
CTR 전결검색 HTML 빌더
======================
엑셀 파일 -> 검색 가능한 HTML(AES-CBC 암호화) 생성

사용법(로컬):
    PASSWORD=비밀번호 python build.py

사용법(GitHub Actions):
    secrets.PAGE_PASSWORD 를 PASSWORD 환경변수로 주입

입력:
    data/*.xlsx 중 가장 최근 파일 (또는 data/source.xlsx)
    template/wrapper.html

출력:
    CTR_전결검색.html  (래퍼 + 새 암호문)

요구 패키지:
    pip install cryptography openpyxl
"""
from __future__ import annotations

import base64
import glob
import html
import json
import os
import sys
from pathlib import Path

from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from openpyxl import load_workbook

ROOT = Path(__file__).parent
DATA_DIR = ROOT / "data"
TEMPLATE = ROOT / "template" / "wrapper.html"
OUTPUT = ROOT / "CTR_전결검색.html"

ITERATIONS = 600_000  # 래퍼의 D.iterations 와 반드시 동일


def pick_excel() -> Path | None:
    """data/ 에서 .xlsx 1개를 골라 반환. source.xlsx 우선, 없으면 mtime 최신. 없으면 None."""
    fixed = DATA_DIR / "source.xlsx"
    if fixed.exists():
        return fixed
    candidates = sorted(
        glob.glob(str(DATA_DIR / "*.xlsx")),
        key=os.path.getmtime,
        reverse=True,
    )
    if not candidates:
        return None
    return Path(candidates[0])


def read_excel(path: Path) -> tuple[list[str], list[list[str]]]:
    """엑셀의 첫 시트를 읽어 (헤더, 행 목록) 반환. 모든 값은 문자열로 변환."""
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    header = [str(c).strip() if c is not None else "" for c in next(rows)]
    body: list[list[str]] = []
    for row in rows:
        # 헤더 길이에 맞춰 자르고, None은 빈 문자열
        cells = [("" if v is None else str(v)) for v in row[: len(header)]]
        # 빈 행 스킵
        if any(c.strip() for c in cells):
            body.append(cells)
    wb.close()
    return header, body


def build_inner_html(header: list[str], rows: list[list[str]], excel_name: str) -> str:
    """복호화 후 화면에 표시될 검색 페이지 HTML 생성."""
    rows_json = json.dumps(rows, ensure_ascii=False)
    header_json = json.dumps(header, ensure_ascii=False)
    excel_name_safe = html.escape(excel_name)
    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=1280">
<title>CTR 전결검색</title>
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:"Pretendard","Apple SD Gothic Neo","Malgun Gothic",system-ui,sans-serif;background:#F4F6FA;color:#1A1F26;padding:24px}}
.hd{{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px}}
.hd .ttl{{font-size:22px;font-weight:700;color:#0B2C5C}}
.hd .meta{{font-size:12px;color:#6B7480}}
.bar{{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px}}
.bar input[type=text]{{flex:1;min-width:240px;padding:10px 14px;border:1.5px solid #DCE1E8;border-radius:8px;font-size:14px;font-family:inherit;outline:none}}
.bar input[type=text]:focus{{border-color:#13427F;box-shadow:0 0 0 3px rgba(19,66,127,.1)}}
.bar select{{padding:10px 12px;border:1.5px solid #DCE1E8;border-radius:8px;font-size:14px;font-family:inherit;background:#fff}}
.cnt{{font-size:13px;color:#6B7480;margin-bottom:8px}}
.tblw{{background:#fff;border:1px solid #E2E6EE;border-radius:10px;overflow:auto;max-height:calc(100vh - 200px)}}
table{{width:100%;border-collapse:collapse;font-size:13px}}
thead th{{position:sticky;top:0;background:#0B2C5C;color:#fff;font-weight:600;padding:10px 12px;text-align:left;white-space:nowrap}}
tbody td{{padding:9px 12px;border-bottom:1px solid #EEF1F6;vertical-align:top}}
tbody tr:hover{{background:#F8FAFD}}
tbody tr.hit mark{{background:#FFF3B0;padding:0 2px;border-radius:2px}}
.empty{{padding:40px;text-align:center;color:#6B7480}}
.btnw{{display:flex;gap:6px}}
.btn{{padding:9px 14px;border:1px solid #DCE1E8;border-radius:8px;background:#fff;font-size:13px;font-family:inherit;cursor:pointer}}
.btn:hover{{border-color:#13427F;color:#13427F}}
</style>
</head>
<body>
<div class="hd">
  <div class="ttl">CTR 전결검색</div>
  <div class="meta">데이터 출처: {excel_name_safe}</div>
</div>
<div class="bar">
  <input id="q" type="text" placeholder="키워드 검색 (여러 단어는 공백으로 구분)">
  <select id="col">
    <option value="-1">전체 컬럼</option>
  </select>
  <div class="btnw">
    <button class="btn" id="clr" type="button">초기화</button>
  </div>
</div>
<div class="cnt" id="cnt"></div>
<div class="tblw">
  <table>
    <thead><tr id="thr"></tr></thead>
    <tbody id="tb"></tbody>
  </table>
</div>
<script>
var H={header_json};
var R={rows_json};
var qEl=document.getElementById("q");
var colEl=document.getElementById("col");
var tbEl=document.getElementById("tb");
var thr=document.getElementById("thr");
var cnt=document.getElementById("cnt");
H.forEach(function(h,i){{
  var th=document.createElement("th");th.textContent=h;thr.appendChild(th);
  var op=document.createElement("option");op.value=i;op.textContent=h;colEl.appendChild(op);
}});
function esc(s){{return s.replace(/[&<>"']/g,function(c){{return{{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[c]}})}}
function mark(s,terms){{
  if(!terms.length) return esc(s);
  var out=esc(s);
  terms.forEach(function(t){{
    if(!t) return;
    var re=new RegExp("("+t.replace(/[.*+?^${{}}()|[\\]\\\\]/g,"\\\\$&")+")","gi");
    out=out.replace(re,"<mark>$1</mark>");
  }});
  return out;
}}
function render(){{
  var q=qEl.value.trim();
  var terms=q.split(/\\s+/).filter(Boolean).map(function(t){{return t.toLowerCase()}});
  var ci=parseInt(colEl.value,10);
  var frag=document.createDocumentFragment();
  var n=0;
  for(var i=0;i<R.length;i++){{
    var row=R[i];
    var hay = ci<0 ? row.join("\\n").toLowerCase() : (row[ci]||"").toLowerCase();
    var ok=terms.every(function(t){{return hay.indexOf(t)>=0}});
    if(!ok) continue;
    n++;
    var tr=document.createElement("tr");
    if(terms.length) tr.className="hit";
    for(var j=0;j<H.length;j++){{
      var td=document.createElement("td");
      td.innerHTML=mark(row[j]||"",terms);
      tr.appendChild(td);
    }}
    frag.appendChild(tr);
    if(n>=2000) break;
  }}
  tbEl.innerHTML="";
  tbEl.appendChild(frag);
  cnt.textContent= R.length+" 건 중 "+n+" 건"+(n>=2000?" (상위 2000건만 표시)":"");
  if(n===0){{
    tbEl.innerHTML='<tr><td colspan="'+H.length+'" class="empty">검색 결과가 없습니다.</td></tr>';
  }}
}}
qEl.addEventListener("input",render);
colEl.addEventListener("change",render);
document.getElementById("clr").addEventListener("click",function(){{qEl.value="";colEl.value="-1";render();qEl.focus();}});
render();
qEl.focus();
</script>
</body>
</html>
"""


def encrypt(plaintext: str, password: str) -> dict:
    """PBKDF2(SHA-256, 600000) + AES-256-CBC + PKCS7."""
    salt = os.urandom(16)
    iv = os.urandom(16)
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=ITERATIONS,
    )
    key = kdf.derive(password.encode("utf-8"))

    data = plaintext.encode("utf-8")
    # PKCS7 padding
    pad_len = 16 - (len(data) % 16)
    data += bytes([pad_len]) * pad_len

    cipher = Cipher(algorithms.AES(key), modes.CBC(iv))
    enc = cipher.encryptor()
    ct = enc.update(data) + enc.finalize()

    return {
        "salt": base64.b64encode(salt).decode("ascii"),
        "iv": base64.b64encode(iv).decode("ascii"),
        "ciphertext": base64.b64encode(ct).decode("ascii"),
        "iterations": ITERATIONS,
    }


def main():
    password = os.environ.get("PASSWORD")
    if not password:
        sys.exit("ERROR: 환경변수 PASSWORD 가 설정되어 있지 않습니다.")

    if not TEMPLATE.exists():
        sys.exit(f"ERROR: 래퍼 템플릿이 없습니다: {TEMPLATE}")

    xlsx = pick_excel()
    if xlsx is None:
        print("ℹ️ data/ 폴더에 엑셀(.xlsx)이 없어 빌드를 건너뜁니다.")
        print("   data/source.xlsx 를 업로드하면 자동으로 빌드됩니다.")
        return
    print(f"[1/4] 엑셀 로드: {xlsx.name}")
    header, rows = read_excel(xlsx)
    print(f"      헤더 {len(header)}개, 데이터 {len(rows)}행")

    print("[2/4] 내부 HTML 생성")
    inner = build_inner_html(header, rows, xlsx.name)
    print(f"      내부 HTML 크기: {len(inner):,} bytes")

    print("[3/4] AES 암호화")
    d = encrypt(inner, password)
    print(f"      ciphertext 크기: {len(d['ciphertext']):,} chars (base64)")

    print("[4/4] 래퍼에 주입 후 저장")
    wrapper = TEMPLATE.read_text(encoding="utf-8")
    d_literal = json.dumps(d, ensure_ascii=False)
    final = wrapper.replace("__D_PLACEHOLDER__", d_literal)
    OUTPUT.write_text(final, encoding="utf-8")
    print(f"  -> {OUTPUT}  ({OUTPUT.stat().st_size:,} bytes)")
    print("완료.")


if __name__ == "__main__":
    main()
