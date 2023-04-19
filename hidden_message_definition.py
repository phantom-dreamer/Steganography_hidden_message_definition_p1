import docx
import MTK2
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_COLOR_INDEX

doc = docx.Document('output.docx')


def run_get_spacing(run):
    rPr = run._r.get_or_add_rPr()
    spacings = rPr.xpath("./w:spacing")
    return spacings


def run_get_scale(run):
    rPr = run._r.get_or_add_rPr()
    scale = rPr.xpath("./w:w")
    return scale


def main():
    code = ''
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            font_color = run.font.color.rgb
            font_size = run.font.size
            font_highlight_color = run.font.highlight_color
            font_scale = run_get_scale(run)
            font_spacing = run_get_spacing(run)

            if (font_color != RGBColor(0, 0, 0) or
                    font_size.pt != 12.0 or
                    font_highlight_color != WD_COLOR_INDEX.WHITE or
                    font_spacing or font_scale):
                for i in range(len(run.text)):
                    code += '1'
                    #print('coooo', code, 'i-', i)
            else:
                for i in range(len(run.text)):
                    code += '0'


    print(f'codes length {len(code)}')
    if len(code) % 16 != 0:
        code += "0" * (16 - len(code) % 16)
        print(f'length after additional value - {len(code)}')
    print(f'code - {code}')

    check = MTK2.MTK2_decode('бог сделал людей, кольт сделал их равными')
    choose = input('enter decoding number \n 1-MTK2 \n 2-koi8_r \n 3-cp866 \n 4-cp1251 \n 5-all types \n')
    match choose:
        case '1':
            plain_text = MTK2.MTK2_decode(code)
            print(f'result with code Bodo / MTK2:\n{plain_text}')
        case '2':
            plain_text = bytes.fromhex(hex(int(code, 2))[2:]).decode(encoding="koi8_r")
            print(f'result with koi8-r: {plain_text}')
        case '3':
            plain_text = bytes.fromhex(hex(int(code, 2))[2:]).decode(encoding="cp866")
            print(f'result with cp866: {plain_text}')
        case '4':
            plain_text = bytes.fromhex(hex(int(code, 2))[2:]).decode(encoding="cp1251")
            print(f'result with cp1251: {plain_text}')
        case '5':
            plain_text = MTK2.MTK2_decode(code)
            print(f'result with code Bodo / MTK2:\n{plain_text}')
            plain_text = bytes.fromhex(hex(int(code, 2))[2:]).decode(encoding="koi8_r")
            print(f'result with koi8-r: {plain_text}')
            plain_text = bytes.fromhex(hex(int(code, 2))[2:]).decode(encoding="cp866")
            print(f'result with cp866: {plain_text}')
            plain_text = bytes.fromhex(hex(int(code, 2))[2:]).decode(encoding="cp1251")
            print(f'result with cp1251: {plain_text}')

if __name__ == '__main__':
    main()
