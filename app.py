from flask import Flask, send_from_directory
import os

from PIL import Image
import requests
import shutil
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage
from urllib.parse import urlparse
from flask import request


def get_filename(url_str):
    url = urlparse(url_str)
    i = len(url.path) - 1
    while i > 0:
        if url.path[i] == '/':
            break
        i = i - 1
    filename = url.path[i + 1:len(url.path)]
    if not filename.strip():
        return False
    return filename


def path_create(path):
    if not os.path.exists(path):
        print("Selected folder not exist, try to create it.")
        os.makedirs(path)
    return


def find_url_img(url):
    path_create('./image')
    filename = './image/' + get_filename(url)
    im = Image.open(requests.get(url, stream=True).raw)
    im.save(filename)
    # urlretrieve(
    #     url,
    #     filename)
    return filename


def todo(data, file_name):
    tpl = DocxTemplate("tzxc.docx")  # 选定模板context = sys.argv[0]  # 需要替换的内容
    context = data
    for street in context['listStreet']:
        for area in street['listArea']:
            area['listProblemAreaNum'] = len(area['listProblemArea'])
            area['listProblemBucketNum'] = len(area['listProblemBucket'])
            if area['listProblemAreaNum'] > 0:
                for problem in area['listProblemArea']:
                    for imgItem in problem['listImg']:
                        imgItem['img'] = InlineImage(tpl, find_url_img(
                            imgItem['picture']), width=Mm(16), height=Mm(16))

    tpl.render(context)  # 渲染替换
    tpl.save('static/' + file_name)  # 保存
    shutil.rmtree('./image')
    return file_name


app = Flask(__name__)


@app.route('/')
def hello_world():
    return 'Hello World!'


@app.route('/docx-templeate', methods=['post'])
def docx():
    data = request.json['data']
    file_name = request.json['file_name']
    todo(data, file_name)
    dirpath = os.path.join(app.root_path, 'static')
    return send_from_directory(
        dirpath,
        file_name,
        as_attachment=True)


if __name__ == '__main__':
    app.run()
