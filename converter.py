import os
import docx


def list_files() -> list:
    paths = []
    for root, dirs, files in os.walk(dir_path):
        if len(files) >= 2:
            set_output_path(output_path)
        for file in files:
            if file.lower().endswith('.docx'):
                paths.append(os.path.join(root, file))
    return paths

def set_output_path(path: str):
    if not os.path.exists(path):
        os.makedirs(path)

def output_result_text_file(index: int, text_list: list):
    receive_file = os.path.join(output_path, '受文者_{0}.txt'.format(index))
    address_file = os.path.join(output_path, '受文者地址_{0}.txt'.format(index))
    for i in range(len(text_list)):
        if i % 2 == 0:
            with open(receive_file, 'a', encoding = 'ANSI') as f1:
                f1.write(text_list[i] + '\n')
        else:
            with open(address_file, 'a', encoding = 'ANSI') as f2:
                f2.write(text_list[i] + '\n')

dir_path = os.path.dirname(os.path.realpath(__file__))
output_path = os.path.join(dir_path, 'output')

for index, file in enumerate(list_files()):
    word_file = docx.Document(file)
    for para in word_file.paragraphs:
        text = para.text
        if ':' not in text and '：' not in text:
            para_list = para.text.split()
            output_result_text_file(index, para_list)
