#!/usr/bin/env python3
"""
Генерирует файл chapters.txt для ffmpeg из .pptx файла.
Использование: python pptx_to_chapters.py presentation.pptx
Затем: ffmpeg -i input.mp4 -i chapters.txt -map_metadata 1 -codec copy output.mp4
"""

import sys
import zipfile
import xml.etree.ElementTree as ET

def parse_pptx_timings(pptx_path):
    timings = []
    
    with zipfile.ZipFile(pptx_path, 'r') as z:
        # Получаем порядок слайдов из presentation.xml
        with z.open('ppt/presentation.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
        
        ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
              'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
        
        # Находим ссылки на слайды по порядку
        slide_refs = root.findall('.//p:sldIdLst/p:sldId', ns)
        
        # Получаем relationship файл чтобы узнать порядок слайдов
        with z.open('ppt/_rels/presentation.xml.rels') as f:
            rels_tree = ET.parse(f)
            rels_root = rels_tree.getroot()
        
        rel_ns = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
        rels = {}
        for rel in rels_root.findall('r:Relationship', rel_ns):
            rels[rel.get('Id')] = rel.get('Target')
        
        # Сопоставляем id → путь к слайду
        slide_paths = []
        for sld in slide_refs:
            rid = sld.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            if rid and rid in rels:
                slide_paths.append('ppt/' + rels[rid].lstrip('/').replace('ppt/', ''))
        
        # Для каждого слайда читаем длительность из transition/timing
        for i, slide_path in enumerate(slide_paths):
            try:
                with z.open(slide_path) as f:
                    stree = ET.parse(f)
                    sroot = stree.getroot()
                
                duration_ms = None
                
                # Ищем длительность в transition (advTm — автопереход в мс)
                trans = sroot.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}transition')
                if trans is not None:
                    adv = trans.get('advTm')
                    if adv is not None:
                        duration_ms = int(adv)
                
                # Если нет advTm — ищем в timing/tnLst (для слайдов с видео и сложной анимацией)
                if duration_ms is None:
                    timing = sroot.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}timing')
                    if timing is not None:
                        # Ищем dur в корневой последовательности
                        for elem in timing.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}seq'):
                            dur = elem.get('dur')
                            if dur and dur != 'indefinite':
                                try:
                                    duration_ms = int(dur)
                                    break
                                except ValueError:
                                    pass
                
                timings.append({
                    'slide': i + 1,
                    'duration_ms': duration_ms,
                    'path': slide_path
                })
                
            except Exception as e:
                print(f"  Предупреждение: слайд {i+1} не удалось прочитать: {e}")
                timings.append({'slide': i + 1, 'duration_ms': None, 'path': slide_path})
    
    return timings


def generate_chapters(timings):
    chapters = []
    current_ms = 0
    
    missing = [t['slide'] for t in timings if t['duration_ms'] is None]
    if missing:
        print(f"\n⚠️  Слайды без автотайминга: {missing}")
        print("   Для них нужно задать длительность вручную (см. ниже).\n")
    
    for t in timings:
        dur = t['duration_ms']
        if dur is None:
            dur = 5000  # дефолт 5 секунд если нет тайминга
        chapters.append({
            'title': f"Слайд {t['slide']}",
            'start_ms': current_ms,
            'end_ms': current_ms + dur - 1,
        })
        current_ms += dur
    
    return chapters


def write_ffmetadata(chapters, output_path='chapters.txt'):
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(';FFMETADATA1\n\n')
        for ch in chapters:
            f.write('[CHAPTER]\n')
            f.write('TIMEBASE=1/1000\n')
            f.write(f"START={ch['start_ms']}\n")
            f.write(f"END={ch['end_ms']}\n")
            f.write(f"title={ch['title']}\n")
            f.write('\n')
    print(f"✅ Готово: {output_path} ({len(chapters)} глав)")
    print(f"   Общая длительность: {chapters[-1]['end_ms'] / 1000:.1f} сек")
    print(f"\nЗапусти ffmpeg:")
    print(f"  ffmpeg -i input.mp4 -i {output_path} -map_metadata 1 -codec copy output.mp4")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Использование: python pptx_to_chapters.py presentation.pptx [output_chapters.txt]")
        sys.exit(1)
    
    pptx_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else 'chapters.txt'
    
    print(f"Читаю {pptx_path}...")
    timings = parse_pptx_timings(pptx_path)
    print(f"Найдено слайдов: {len(timings)}")
    
    chapters = generate_chapters(timings)
    write_ffmetadata(chapters, output_path)