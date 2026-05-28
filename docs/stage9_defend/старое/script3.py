#!/usr/bin/env python3
"""
Извлекает тайминги слайдов через PowerPoint COM (только Windows с установленным PPT).
Использование: python pptx_com_timings.py presentation.pptx
"""

import sys
import os

try:
    import win32com.client
except ImportError:
    print("Нужен pywin32: pip install pywin32")
    sys.exit(1)

def get_timings_via_com(pptx_path):
    pptx_path = os.path.abspath(pptx_path)
    
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Visible = True  # False иногда глючит с медиа

    print(f"Открываю {pptx_path}...")
    pres = ppt.Presentations.Open(pptx_path, ReadOnly=True, WithWindow=False)

    timings = []
    total_slides = pres.Slides.Count
    print(f"Слайдов: {total_slides}")

    for i in range(1, total_slides + 1):
        slide = pres.Slides(i)
        
        # SlideShowTransition.AdvanceTime — время автоперехода в секундах
        # SlideShowTransition.AdvanceOnTime — включён ли автопереход
        trans = slide.SlideShowTransition
        advance_on_time = trans.AdvanceOnTime
        advance_time = trans.AdvanceTime  # секунды (float)

        # Если автопереход не задан — берём длительность через Shape (видео)
        duration_sec = None
        if advance_on_time and advance_time > 0:
            duration_sec = advance_time
        else:
            # Ищем медиа-объекты на слайде и берём максимальную длительность
            for j in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes(j)
                try:
                    # ppMediaTypeMovie = 3, ppMediaTypeSound = 2
                    if shape.MediaType in (2, 3):
                        media_dur = shape.MediaFormat.Duration  # секунды
                        if media_dur and media_dur > 0:
                            if duration_sec is None or media_dur > duration_sec:
                                duration_sec = media_dur
                except Exception:
                    pass

        timings.append({
            'slide': i,
            'duration_sec': duration_sec,
            'advance_on_time': advance_on_time,
        })
        
        status = f"{duration_sec:.2f}s" if duration_sec else "⚠️  нет тайминга"
        print(f"  Слайд {i:3d}: {status}")

    pres.Close()
    ppt.Quit()
    return timings


def write_ffmetadata(timings, output_path='chapters.txt', default_sec=5.0):
    missing = [t['slide'] for t in timings if t['duration_sec'] is None]
    if missing:
        print(f"\n⚠️  Слайды без тайминга (поставлено {default_sec}с по дефолту): {missing}")

    chapters = []
    current_ms = 0
    for t in timings:
        dur_ms = int((t['duration_sec'] or default_sec) * 1000)
        chapters.append({
            'title': f"Слайд {t['slide']}",
            'start_ms': current_ms,
            'end_ms': current_ms + dur_ms - 1,
        })
        current_ms += dur_ms

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(';FFMETADATA1\n\n')
        for ch in chapters:
            f.write('[CHAPTER]\n')
            f.write('TIMEBASE=1/1000\n')
            f.write(f"START={ch['start_ms']}\n")
            f.write(f"END={ch['end_ms']}\n")
            f.write(f"title={ch['title']}\n\n")

    print(f"\n✅ Готово: {output_path} ({len(chapters)} глав)")
    print(f"   Общая длительность: {chapters[-1]['end_ms'] / 1000:.1f} сек")
    print(f"\nЗапусти ffmpeg:")
    print(f"  ffmpeg -i input.mp4 -i {output_path} -map_metadata 1 -codec copy output.mp4")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Использование: python pptx_com_timings.py presentation.pptx [chapters.txt]")
        sys.exit(1)

    pptx_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else 'chapters.txt'

    timings = get_timings_via_com(pptx_path)
    write_ffmetadata(timings, output_path)