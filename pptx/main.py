import api_key
import lyricsgenius as lg
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt, Cm
import platform
import getpass
import os


artist_input = input("Specify artist:")
artist_song = input("Specify song:")
def get_lyrics():
        genius = lg.Genius(api_key.client_access_token, skip_non_songs=True, excluded_terms=["(Remix)", "(Live)"], remove_section_headers=True)
        artist = genius.search_artist(artist_input, max_songs=0, sort="title")
        song = artist.song(artist_song)
        song_lyrics = song.lyrics
        remove_embed = song_lyrics.replace("Embed", '')
        remove_digits = ''.join([i for i in remove_embed if not i.isdigit()])
        split_lyrics = remove_digits.splitlines()
        lyrics = [x for x in split_lyrics if x]
        return lyrics

def make_pres():
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
    for i in get_lyrics():
        slide=pres.slides.add_slide(layout)
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        txBox.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        tf.word_wrap = True
        p = tf.add_paragraph()
        p.text = i
        p.font.size = Pt(60)
        if platform.system() == 'Windows':
            pres.save(f"C:/Users/{getpass.getuser()}/Desktop/{directory}/{artist_song} by {artist_input}.pptx")
        elif platform.system() == 'Darwin':
            pres.save(f"/Users/{getpass.getuser()}/Desktop/{directory}/{artist_song} by {artist_input}.pptx")
        elif platform.system() == 'Linux':
            pres.save(f"/home/{getpass.getuser()}/Desktop/{directory}/{artist_song} by {artist_input}.pptx")
        else:
            print("Unknown OS...")

def main():
    make_pres()

if __name__ == '__main__':
    raise SystemExit(main())
