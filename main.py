#!/usr/bin/env python3

import ffmpeg
import click
import openpyxl
import datetime
import platform

@click.group()
def luogu_tools():
    pass

def get_col_time(col_st, col_ed):

    for st, ed in map(lambda x, y: (x.value, y.value), col_st[1:], col_ed[1:]):

        if type(st) == datetime.datetime:
            st = st.time()
        if type(ed) == datetime.datetime:
            ed = ed.time()

        if st is None and ed is None:
            break
        elif st is not None and ed is not None:
            yield st, ed
        else:
            click.echo('Wrong format at row{}.'.format(st.row))

@click.command()
@click.option('--excel',
        help = 'excel file of the course',
        type = click.Path(exists = True, file_okay = True, readable = True),
        required = True
)
@click.option('--mp3',
        help = 'mp3 file of the course',
        type = click.Path(exists = True, file_okay = True, readable = True),
        required = True
)
def lg_audio(excel, mp3):
    try:
        xls = openpyxl.load_workbook(filename = excel)
    except:
        print('Something went wrong about the ' + excel)
        return

    ffmpeg_params = {'format': 'mp3'}
    if platform.system() == 'Darwin':
        ffmpeg_params['c:v'] = 'h264_videotoolbox';

    sheet = xls.worksheets[0]
    time_st, time_ed = list(sheet.columns)[1:3]

    for index, (st, ed) in enumerate(get_col_time(time_st, time_ed)):
        print(index, '-ss {} -t {}'.format(st, ed))
        audio = ffmpeg.input(mp3).audio
        audio = audio.filter('atrim', start = st, end = ed)
        out = ffmpeg.output(audio, '{}.mp3'.format(index + 1),  format = 'mp3')
        out.run()
        
        '''(
            ffmpeg
            .input(mp3)
            .trim(start = st, end = ed)
            .output('{}.mp3'.format(index + 1))
            .run()
        )'''

luogu_tools.add_command(lg_audio)
if __name__ == '__main__':
    lg_audio()
