"""Create a project without overwriting an existing project."""
import argparse
import json
import re
import shutil
from datetime import datetime, timedelta, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('name', help='작업 이름')
    parser.add_argument('--title', help='발표 제목')
    parser.add_argument('--folder', help='날짜 접두사 없이 사용할 폴더 이름')
    args = parser.parse_args()
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', args.name).strip(' .')
    reserved={'CON','PRN','AUX','NUL'} | {f'COM{i}' for i in range(1,10)} | {f'LPT{i}' for i in range(1,10)}
    if not name or name in {'.','..'} or name.split('.')[0].upper() in reserved:
        parser.error('유효한 작업 이름을 입력하세요.')
    day = datetime.now(timezone(timedelta(hours=9))).strftime('%Y%m%d')
    folder = args.folder or f'{day}_{name}'
    if folder in {'.','..'} or folder.split('.')[0].upper() in reserved or re.search(r'[<>:"/\\|?*\x00-\x1f]', folder) or folder.endswith((' ','.')):
        parser.error('folder는 경로가 아닌 하나의 폴더 이름이어야 합니다.')
    paths = {kind: ROOT/kind/folder for kind in ('inputs','ppt_info','output')}
    if any(p.exists() for p in paths.values()):
        parser.error('같은 작업 폴더가 이미 있습니다. 다른 이름을 사용하세요.')
    template = ROOT/'templet/base.pptx'
    if not template.is_file():
        parser.error('templet/base.pptx가 없습니다.')
    for path in paths.values():
        path.mkdir(parents=True)
    title = args.title or args.name
    brief = (f'# {title}\n\n'
             '## 목적과 청중\n\n- 목적:\n- 청중:\n- 발표 시간 또는 분량:\n\n'
             '## 내용과 자료\n\n- 핵심 내용:\n- 자료 파일 또는 링크:\n- 반드시 포함할 사항:\n\n'
             '## 이번 작업의 추가 요청\n\n- 구성·디자인:\n- 수치 표시 방식:\n- 기타:\n\n'
             f'## 결과 위치\n\n`output/{folder}/`에 PPTX와 PDF를 저장합니다.\n')
    (paths['inputs']/'brief.md').write_text(brief,encoding='utf8')
    shutil.copy2(template,paths['output']/(name+'.pptx'))
    record = dict(title=title, folder=folder, presentation=name+'.pptx')
    (paths['ppt_info']/'project.json').write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding='utf8')
    print(f'Created project: {folder}')
    print(f'Brief: inputs/{folder}/brief.md')
    print(f'Presentation: output/{folder}/{name}.pptx')
    print('템플릿 예문을 실제 내용으로 바꾸고, 필요 없는 슬라이드를 삭제하세요.')

if __name__ == '__main__':
    main()
