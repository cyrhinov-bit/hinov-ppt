from PIL import Image, ImageFilter
import glob

targets = [
    'C:/Users/GNONSKAN EVARISTE/Desktop/hinov ppt/assets/img_hinov_main.jpeg',
    'C:/Users/GNONSKAN EVARISTE/Desktop/hinov ppt/assets/img_vision.jpeg',
    'C:/Users/GNONSKAN EVARISTE/Desktop/hinov ppt/assets/img_mission.jpeg',
    'C:/Users/GNONSKAN EVARISTE/Desktop/hinov ppt/assets/img_expertise.jpeg',
    'C:/Users/GNONSKAN EVARISTE/Desktop/hinov ppt/assets/img_engagement.jpeg',
]

for p in targets:
    im = Image.open(p).convert('RGB')
    w, h = im.size
    im = im.resize((int(w * 2), int(h * 2)), Image.LANCZOS)
    im = im.filter(ImageFilter.UnsharpMask(radius=2, percent=180))
    if p.lower().endswith('.jpg') or p.lower().endswith('.jpeg'):
        im.save(p, quality=95)
    else:
        im.save(p, optimize=True)
    print('SHARPENED/ UPSCALED', p, (w, h), '->', im.size)

for p in sorted(glob.glob('C:/Users/GNONSKAN EVARISTE/Desktop/hinov ppt/assets/partner_*.jp*g')):
    im = Image.open(p).convert('RGB')
    w, h = im.size
    im = im.resize((int(w * 2), int(h * 2)), Image.LANCZOS)
    im = im.filter(ImageFilter.UnsharpMask(radius=2, percent=150))
    im.save(p, optimize=True)
    print('SHARPENED/ UPSCALED', p, (w, h), '->', im.size)

print('DONE')
