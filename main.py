import os

from docx import Document
from docx.shared import Cm
from docxcompose.composer import Composer
from docxtpl import DocxTemplate, InlineImage


def dowork(filepath):
    allfiles = os.listdir(filepath)
    # 遍历该文件夹下的所有文件夹
    for file in allfiles:
        tem_path = os.path.join(filepath, file)
        # print(file == '三改一拆')
        # 判断该文件夹类型
        if file == '三改一拆':
            print(file)
            # 遍历得到所有年份目录
            for year_path in getdir(tem_path):
                print('获取年份目录 ：' + year_path)
                docname = year_path + '年' + file + '台账' + '.docx'
                document = Document()
                target_composer = Composer(document)
                # 遍历年份目录得到所有月份目录
                for month_path in getdir(os.path.join(filepath, file, year_path)):
                    print('获取月份目录 ：' + month_path)
                    # 遍历所有月份目录，得到每月台账资料
                    for mc_path in getdir(os.path.join(filepath, file, year_path, month_path)):
                        print('获取每月台账资料 ：' + mc_path)
                        piclist = []
                        txtlist = []
                        for mc_file in os.listdir(os.path.join(filepath, file, year_path, month_path, mc_path)):
                            print('获取每项台账资料文件 ：' + mc_file)
                            if os.path.splitext(mc_file)[1] == '.txt':
                                txtlist.append(mc_file)
                            else:
                                piclist.append(os.path.join(filepath, file, year_path, month_path, mc_path, mc_file))
                        if txtlist.__len__() > 1:
                            print('there are more than one txt file in ' + os.path.join(filepath, file, year_path,
                                                                                        month_path, mc_path))
                        else:
                            tem_docx = DocxTemplate("template1.docx")
                            result_content_dict = {}
                            count = 0
                            while count < piclist.__len__():
                                image = InlineImage(tem_docx, piclist[count], width=Cm(7), height=Cm(5))
                                result_content_dict["pic" + str(count)] = image
                                count += 1
                            for line in open(os.path.join(filepath, file, year_path, month_path, mc_path, txtlist[0]),
                                             "r", encoding="utf-8"):  # 设置文件对象并读取每一行文件
                                line = line.strip("\n")
                                data_list = line.split("#", 1)
                                if len(data_list):
                                    result_content_dict[data_list[0]] = data_list[1]
                            print(result_content_dict)
                            tem_docx.render(result_content_dict)
                            tem_docx.add_page_break()
                            target_composer.append(tem_docx)
                target_composer.save(os.path.join(filepath, file, year_path) + "\\" + docname)
        elif file == '安全生产':
            print(file)
        else:
            print("无此类型台账模板：" + file)
        # if os.path.isdir(tem_path):
        #     getDir(tem_path)
        # 如果不是文件夹，保存文件路径及文件名
        # elif os.path.isfile(tem_path):
        #     allpath.append(tem_path)
        #     allname.append(file)


def getdir(filepath):
    allpath = []
    allfiles = os.listdir(filepath)
    # 遍历该文件夹下的所有文件夹
    for file in allfiles:
        tem_path = os.path.join(filepath, file)
        if os.path.isdir(tem_path):
            allpath.append(file)
    return allpath


def getcontent(txtfile):
    dict = {}
    for line in open(txtfile, "r", encoding="utf-8"):  # 设置文件对象并读取每一行文件
        line = line.strip("\n")
        datalist = line.split("#", 1)
        if len(datalist):
            dict[datalist[0]] = datalist[1]
    return dict


if __name__ == "__main__":
    path = "D:\\testpackage"
    folder = os.path.exists(path)
    if not folder:
        print("路径不存在")
    else:
        dowork(path)
    # for file in allpath:
    #     print(file)
    # print("-------------------------")
    # for name in allname:
    #     print(name)
