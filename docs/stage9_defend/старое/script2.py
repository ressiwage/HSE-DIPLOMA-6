#!/usr/bin/env python3
"""
Дампит XML первых N слайдов для диагностики тайминга.
Использование: python diagnose_pptx.py presentation.pptx
"""

import sys
import zipfile
import xml.etree.ElementTree as ET

def diagnose(pptx_path, num_slides=2):
    with zipfile.ZipFile(pptx_path, 'r') as z:
        # Список всех файлов в архиве
        all_files = z.namelist()
        slide_files = sorted([f for f in all_files if f.startswith('ppt/slides/slide') and '_rels' not in f])
        
        print(f"Всего слайдов найдено: {len(slide_files)}")
        print(f"Все файлы в ppt/slides/: {[f for f in all_files if 'slides' in f][:20]}\n")
        
        for slide_path in slide_files[:num_slides]:
            print(f"\n{'='*60}")
            print(f"СЛАЙД: {slide_path}")
            print('='*60)
            with z.open(slide_path) as f:
                content = f.read().decode('utf-8')
            
            # Ищем секцию timing
            if '<p:timing' in content:
                start = content.find('<p:timing')
                end = content.find('</p:timing>') + len('</p:timing>')
                print("TIMING БЛОК НАЙДЕН:")
                print(content[start:end][:3000])
            else:
                print("⚠️  Блок <p:timing> НЕ НАЙДЕН на этом слайде")
            
            # Ищем transition
            if '<p:transition' in content:
                start = content.find('<p:transition')
                end = content.find('>', start) + 1
                # Захватим до закрывающего тега
                end2 = content.find('</p:transition>')
                if end2 > 0:
                    end = end2 + len('</p:transition>')
                print("\nTRANSITION БЛОК:")
                print(content[start:end])
            else:
                print("\n⚠️  Блок <p:transition> НЕ НАЙДЕН")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Использование: python diagnose_pptx.py presentation.pptx")
        sys.exit(1)
    diagnose(sys.argv[1])