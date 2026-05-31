"""마크다운 파일을 깔끔한 흰 배경 HTML 보고서로 변환"""
import sys
import os
from datetime import date

def convert(md_path: str, html_path: str) -> None:
    with open(md_path, encoding="utf-8") as f:
        md_text = f.read()

    try:
        import markdown
        body_html = markdown.markdown(
            md_text,
            extensions=["tables", "fenced_code", "toc"]
        )
    except ImportError:
        # markdown 라이브러리 없으면 기본 변환
        body_html = basic_md_to_html(md_text)

    title = ""
    for line in md_text.splitlines():
        if line.startswith("# "):
            title = line[2:].strip()
            break
    if not title:
        title = os.path.basename(md_path)

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{title}</title>
  <style>
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
      font-family: 'Malgun Gothic', '맑은 고딕', -apple-system, sans-serif;
      font-size: 15px;
      line-height: 1.8;
      color: #222;
      background: #fff;
      max-width: 860px;
      margin: 0 auto;
      padding: 40px 32px;
    }}
    h1 {{ font-size: 26px; font-weight: 700; margin: 0 0 8px; color: #111; border-bottom: 3px solid #2563eb; padding-bottom: 12px; }}
    h2 {{ font-size: 20px; font-weight: 700; margin: 36px 0 12px; color: #1e3a5f; border-left: 4px solid #2563eb; padding-left: 12px; }}
    h3 {{ font-size: 17px; font-weight: 600; margin: 24px 0 8px; color: #2d4a6b; }}
    p {{ margin: 10px 0; }}
    ul, ol {{ margin: 10px 0 10px 24px; }}
    li {{ margin: 5px 0; }}
    a {{ color: #2563eb; text-decoration: none; }}
    a:hover {{ text-decoration: underline; }}
    blockquote {{ border-left: 4px solid #e5e7eb; padding: 8px 16px; color: #555; margin: 16px 0; background: #f9fafb; }}
    table {{ border-collapse: collapse; width: 100%; margin: 16px 0; font-size: 14px; }}
    th {{ background: #2563eb; color: #fff; padding: 10px 14px; text-align: left; }}
    td {{ padding: 9px 14px; border-bottom: 1px solid #e5e7eb; }}
    tr:nth-child(even) td {{ background: #f8faff; }}
    code {{ background: #f3f4f6; padding: 2px 6px; border-radius: 4px; font-size: 13px; font-family: 'Consolas', monospace; }}
    pre {{ background: #f3f4f6; padding: 16px; border-radius: 6px; overflow-x: auto; margin: 16px 0; }}
    pre code {{ background: none; padding: 0; }}
    hr {{ border: none; border-top: 1px solid #e5e7eb; margin: 32px 0; }}
    .meta {{ color: #666; font-size: 13px; margin: 8px 0 32px; }}
    @media print {{
      body {{ max-width: 100%; padding: 20px; }}
      a {{ color: #000; }}
    }}
  </style>
</head>
<body>
{body_html}
<hr>
<p style="color:#aaa;font-size:12px;text-align:right;">생성: {date.today()}</p>
</body>
</html>"""

    os.makedirs(os.path.dirname(html_path) or ".", exist_ok=True)
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"HTML 저장 완료: {html_path}")


def basic_md_to_html(text: str) -> str:
    """markdown 라이브러리 없을 때 사용하는 기본 변환"""
    import re
    lines = text.splitlines()
    result = []
    in_list = False

    for line in lines:
        if line.startswith("# "):
            if in_list: result.append("</ul>"); in_list = False
            result.append(f"<h1>{line[2:].strip()}</h1>")
        elif line.startswith("## "):
            if in_list: result.append("</ul>"); in_list = False
            result.append(f"<h2>{line[3:].strip()}</h2>")
        elif line.startswith("### "):
            if in_list: result.append("</ul>"); in_list = False
            result.append(f"<h3>{line[4:].strip()}</h3>")
        elif line.startswith("- ") or line.startswith("• "):
            if not in_list: result.append("<ul>"); in_list = True
            content = line[2:].strip()
            result.append(f"<li>{inline_md(content)}</li>")
        elif line.strip() == "---":
            if in_list: result.append("</ul>"); in_list = False
            result.append("<hr>")
        elif line.strip() == "":
            if in_list: result.append("</ul>"); in_list = False
            result.append("")
        else:
            if in_list: result.append("</ul>"); in_list = False
            result.append(f"<p>{inline_md(line)}</p>")

    if in_list:
        result.append("</ul>")
    return "\n".join(result)


def inline_md(text: str) -> str:
    import re
    text = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", text)
    text = re.sub(r"\*(.+?)\*", r"<em>\1</em>", text)
    text = re.sub(r"`(.+?)`", r"<code>\1</code>", text)
    text = re.sub(r"\[(.+?)\]\((https?://[^\)]+)\)", r'<a href="\2">\1</a>', text)
    return text


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("사용법: python md_to_html.py <입력.md> <출력.html>")
        sys.exit(1)
    convert(sys.argv[1], sys.argv[2])
