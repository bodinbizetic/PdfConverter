import win32com.client as win32
import os
import sys


def convertDoc(input_file_path, output_file_path):
    word = None
    deck = None
    try:
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = True

        deck = word.Documents.Open(input_file_path)
        deck.SaveAs(output_file_path, 17)  # 17 = WdFormatPdf
    finally:
        if deck is not None:
            deck.Close()
        if word is not None:
            word.Quit()


def convertPPT(input_file_path, output_file_path):
    deck = None
    powerpoint = None
    try:
        powerpoint = win32.gencache.EnsureDispatch('Powerpoint.Application')
        powerpoint.Visible = True
        deck = powerpoint.Presentations.Open(input_file_path)
        deck.SaveAs(output_file_path, 32)  # 32 = PpFormatPdf
    finally:
        if deck is not None:
            deck.Close()
        if powerpoint is not None:
            powerpoint.Quit()


def getOutputFileName(input_path, new_ext) -> str:
    pathName, ext = os.path.splitext(input_path)
    output_file = pathName + new_ext
    i = 1
    while os.path.exists(output_file):
        output_file = pathName + ' (' + str(i) + ')' + new_ext
    return output_file


def getExtension(input_path) -> str:
    return os.path.splitext(input_path)[1]


def getConverter(input_path):
    ext = getExtension(input_path)

    if ext in ['.doc', '.docx']:
        return convertDoc
    elif ext in ['.ppt', '.pptx']:
        return convertPPT
    else:
        raise NotImplemented('File extension not supported')


def convertFile(input_file):
    function = getConverter(input_file)
    output_file = getOutputFileName(input_file, '.pdf')
    function(input_file, output_file)


def main():
    path_msg = []
    error_happen = False
    for path in sys.argv[1:]:
        try:
            if not os.path.exists(path):
                error_happen = True
                path_msg.append((path, 'File not exists'))
                continue
            convertFile(path)
            path_msg.append((path, 'OK'))
        except Exception as e:
            error_happen = True
            path_msg.append((path, str(e.with_traceback())))

    for path, error in path_msg:
        print(path, ': -> ', error)
    
    if error_happen:
        input('Enter key to continue...')


if __name__ == '__main__':
    main()
