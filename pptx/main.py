import api_key
import lyricsgenius as lg
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt, Cm
import platform
import getpass
import os
from gooey import Gooey, GooeyParser

# Mac colors: footer_bg_color="#789CA4", sidebar_bg_czolor="#789CA4", body_bg_color="#789CA4", header_bg_color="#789CA4"
@Gooey(program_name='Auto-pptx', program_description="A simple, intuitive powerpoint creator for all.")
def parse_arguments():
    parser = GooeyParser()
    parser.add_argument('-a', '--Artist', type=str, nargs='+', required=True, help='Specify what artist you want to search.')
    parser.add_argument('-s', '--Song', type=str, nargs='+', required=True, help='Specify what song you want to search.')
    parser.add_argument('--Font', type=str, nargs='+', required=False, help="Specify the font you want")
    parser.add_argument('--Font_Size', type=int, required=False, help="Specify the font size")
    parser.add_argument('--Background_Color', widget="ColourChooser", required=False, help="Choose background color")
    parser.add_argument('--Font_Color', widget="ColourChooser", required=False, help="Choose font color")
    args = parser.parse_args()
    artist_input = args.Artist
    artist_song = args.Song
    font = args.Font
    font_size = args.Font_Size
    background_color_hex = args.Background_Color
    font_color_hex = args.Font_Color
    if background_color_hex != None:
        background_color_strip = background_color_hex.lstrip("#")
        background_color = tuple(int(background_color_strip[i:i+2], 16) for i in (0, 2, 4))
    else:
        background_color = None
    if font_color_hex != None:
        font_color_strip = font_color_hex.lstrip("#")
        font_color = tuple(int(font_color_strip[i:i+2], 16) for i in (0, 2, 4))
    else:
        font_color = None
    return artist_input, artist_song, font, font_size, background_color, font_color

def get_lyrics(artist_input, artist_song, font, font_size, background_color, font_color):
        genius = lg.Genius(api_key.client_access_token, skip_non_songs=True,remove_section_headers=True)
        artist = genius.search_artist(' '.join(artist_input), max_songs=0, sort="title")
        song = artist.song(' '.join(artist_song))
        song_lyrics = song.lyrics
        remove_embed = song_lyrics.replace("Embed", '')
        remove_digits = ''.join([i for i in remove_embed if not i.isdigit()])
        split_lyrics = remove_digits.splitlines()
        lyrics = [x for x in split_lyrics if x]
        return lyrics

def make_pres(artist, song, font, font_size, background_color, font_color):
    pres = Presentation()
    layout = pres.slide_layouts[6]
    left = Cm(3)
    top = Cm(4)
    width = Cm(20)
    height = Cm(5)
    directory = "Powerpoint Lyrics"
    if platform.system() == 'Windows':
        parent_dir = f"C:/Users/{getpass.getuser()}/Desktop"
        path = os.path.join(parent_dir, directory)
        os.makedirs(path, exist_ok=True)
    elif platform.system() == 'Darwin':
        parent_dir = f"/Users/{getpass.getuser()}/Desktop"
        path = os.path.join(parent_dir, directory)
        os.makedirs(path, exist_ok=True)
    elif platform.system() == 'Linux':
        parent_dir = f"/home/{getpass.getuser()}/Desktop"
        path = os.path.join(parent_dir, directory)
        os.makedirs(path, exist_ok=True)
    else:
        print("Unknown OS...can't create directory")
    for i in get_lyrics(*parse_arguments()):
        slide=pres.slides.add_slide(layout)
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        txBox.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        tf.word_wrap = True
        p = tf.add_paragraph()
        p.text = i
        if background_color != None:
            r, g, b = background_color
            background = slide.background
            fill = background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(r, g, b)
        if font_color != None:
            r, g, b = font_color
            p.font.color.rgb = RGBColor(r, g, b)
        if font != None:
            p.font.name = ' '.join(font)
        if font_size != None:
            p.font.size = Pt(font_size)
        else:
            p.font.size = Pt(60)
        if platform.system() == 'Windows':
            pres.save(f"C:/Users/{getpass.getuser()}/Desktop/{directory}/{' '.join(song)} by {' '.join(artist)}.pptx")
        elif platform.system() == 'Darwin':
            pres.save(f"/Users/{getpass.getuser()}/Desktop/{directory}/{' '.join(song)} by {' '.join(artist)}.pptx")
        elif platform.system() == 'Linux':
            pres.save(f"/home/{getpass.getuser()}/Desktop/{directory}/{' '.join(song)} by {' '.join(artist)}.pptx")
        else:
            print("Unknown OS...")

def main():
    make_pres(*parse_arguments())


if __name__ == '__main__':
    raise SystemExit(main())