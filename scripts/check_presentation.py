"""Basic PPTX checks. Passing does not replace visual review."""
import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS={'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p':'http://schemas.openxmlformats.org/presentationml/2006/main'}
EMU_PER_PT=12700

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('presentation',type=Path)
    parser.add_argument('--audit',type=Path)
    parser.add_argument('--line-spacing',type=float,default=1.3)
    parser.add_argument('--allow-placeholders',action='store_true')
    parser.add_argument('--minimum-margin-cm',type=float,default=0)
    args=parser.parse_args()
    errors=[]
    warnings=[]
    with zipfile.ZipFile(args.presentation) as z:
        bad=z.testzip()
        if bad: errors.append(f'Corrupt ZIP member: {bad}')
        root=ET.fromstring(z.read('ppt/presentation.xml'))
        size=root.find('p:sldSz',NS)
        width,height=int(size.get('cx')),int(size.get('cy'))
        files=sorted([n for n in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml',n)],
                     key=lambda n:int(re.search(r'slide(\d+)',n).group(1)))
        if not files: errors.append('No slides.')
        for number,name in enumerate(files,1):
            slide=ET.fromstring(z.read(name))
            for shape in slide.findall('./p:cSld/p:spTree/*',NS):
                transform=shape.find('./p:spPr/a:xfrm',NS)
                if transform is None: transform=shape.find('./p:xfrm',NS)
                if transform is not None:
                    off,extent=transform.find('a:off',NS),transform.find('a:ext',NS)
                    if off is not None and extent is not None:
                        x,y,cx,cy=int(off.get('x')),int(off.get('y')),int(extent.get('cx')),int(extent.get('cy'))
                        tolerance=EMU_PER_PT
                        if x < -tolerance or y < -tolerance or x+cx > width+tolerance or y+cy > height+tolerance:
                            errors.append(f'Slide {number}: object extends outside slide.')
                        is_content=bool(shape.findall('.//a:t',NS)) or shape.tag.endswith(('pic','graphicFrame'))
                        margin=round(args.minimum_margin_cm*360000)
                        if is_content and (x < margin-tolerance or y < margin-tolerance or x+cx > width-margin+tolerance or y+cy > height-margin+tolerance):
                            errors.append(f'Slide {number}: content margin is less than {args.minimum_margin_cm} cm.')
            for p in slide.findall('.//a:p',NS):
                text=''.join(n.text or '' for n in p.findall('.//a:t',NS))
                if not text: continue
                spacing=p.find('./a:pPr/a:lnSpc/a:spcPct',NS)
                if spacing is None:
                    warnings.append(f'Slide {number}: inherited/unspecified line spacing; inspect visually.')
                elif abs(int(spacing.get('val'))-round(args.line_spacing*100000)) > 1:
                    errors.append(f'Slide {number}: line spacing differs from {args.line_spacing}.')
                if not args.allow_placeholders and any(s in text for s in ('[발표 제목]','[항목]','이미지 영역','여기에 입력')):
                    warnings.append(f'Slide {number}: template placeholder may remain.')
    if args.audit:
        audit=json.loads(args.audit.read_text(encoding='utf-8-sig'))
        for shape in audit:
            if shape['boundWidth'] > shape.get('availableWidth',shape['width'])+1 or shape['boundHeight'] > shape.get('availableHeight',shape['height'])+1:
                errors.append(f"Slide {shape['slide']}: possible clipped text in {shape['name']}.")
    report=dict(presentation=str(args.presentation.resolve()),slide_count=len(files),
                slide_size_pt=[round(width/EMU_PER_PT,2),round(height/EMU_PER_PT,2)],
                errors=errors,warnings=sorted(set(warnings)),visual_review_required=True)
    print(json.dumps(report,ensure_ascii=True,indent=2))
    return bool(errors)

if __name__=='__main__':
    sys.exit(main())
