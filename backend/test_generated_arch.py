import sys, os
sys.path.append('.')
from services.visual_engine import _render_arch_matplotlib

xml = open('arch_slide.xml', encoding='utf-8').read()

png = _render_arch_matplotlib(xml)

print('PNG:', png)

if png and os.path.exists(png):
    print('PNG size:', os.path.getsize(png), 'bytes')
    print('SUCCESS - Generated architecture diagram rendered!')
else:
    print('PNG MISSING - Renderer failed!')
