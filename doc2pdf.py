import os
import argparse
import time

import comtypes.client
import logging

# 设置word和pdf文件类型常量
wdFormatPDF = 17
wdFormatDoc = 0
wdFormatDocx = 12


def parse_options():
    parser = argparse.ArgumentParser(description='This is a batch tool for converting Word documents to PDF files.')
    parser.add_argument('-s', '--source', type=str, dest="source", required=True, default='', help="source file path")
    parser.add_argument('-t', '--target', type=str, dest="target", required=False,
                        help="save target pdf file path , default is same as source file path")
    args = parser.parse_args()
    return args


def get_logger(logger_name):
    logger = logging.getLogger(logger_name)
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s %(name)s %(levelname)s: %(message)s')
    handler = logging.StreamHandler()
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    return logger


logger = get_logger("doc2pdf")


def get_all_files(path):
    """
    获取指定目录下的所有文件
    :param path:
    :return:
    """
    files = []
    if os.path.exists(path) and os.path.isdir(path):
        for i in os.scandir(path):
            if i.is_dir():
                files.extend(get_all_files(os.path.join(path, i)))
            if i.is_file():
                files.append(os.path.join(path, i))
    elif os.path.exists(path) and os.path.isfile(path):
        files.append(path)

    return files


def doc2pdf(doc_path, pdf_path):
    """
    将word 转换成pdf
    :param doc_path:
    :param pdf_path:
    :return:
    """

    if doc_path.endswith('.doc') or doc_path.endswith('.docx'):
        try:
            word_app = comtypes.client.CreateObject('Word.Application')
            word_app.Visible = False
            doc = word_app.Documents.Open(doc_path)
            doc.SaveAs(pdf_path, FileFormat=wdFormatPDF)
            doc.Close()
            word_app.Quit()
        except Exception as e:
            return False, e
    else:
        False, 'The file format is not supported, only doc or docx formats are supported'
    return True, ''


if __name__ == "__main__":
    args = parse_options()

    start_time = time.time()
    doc_dir = os.path.abspath(args.source)
    doc_files = get_all_files(doc_dir)

    result = {
        "failed": 0,
        "success": 0,
        "total": 0
    }

    for file in doc_files:
        target_pdf_file_path = os.path.splitext(file)[0] + '.pdf'

        if args.target:
            target_pdf_dir = os.path.dirname(file.replace(os.path.abspath(args.source), os.path.abspath(args.target)))
            target_pdf_file_name = os.path.splitext(os.path.basename(file))[0]
            os.makedirs(target_pdf_dir, exist_ok=True)
            target_pdf_file = os.path.join(target_pdf_dir, target_pdf_file_name + '.pdf')

        if file.endswith('.doc') or file.endswith('.docx'):
            result['total'] += 1

            logger.info(f'Start convert file {file}')
            ret, info = doc2pdf(file, target_pdf_file)

            if not ret:
                logger.error(f'Convert file {file} error: {info}')
                result['failed'] += 1
            else:
                result['success'] += 1
                logger.info(f'Convert file {file} succcess, save pdf to {target_pdf_file}')

    end_time = time.time()
    logger.info(
        f"All finish!, convert {result['total']} files, success {result['success']}, failed {result['failed']}, "
        f"cost time:{int(end_time - start_time)} seconds")
