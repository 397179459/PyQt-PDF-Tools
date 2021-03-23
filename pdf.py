from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt5.uic import loadUi
import sys, os
from PyPDF2 import PdfFileReader, PdfFileWriter


class MainWindow(QMainWindow):
    def __init__(self, *args):
        super(MainWindow, self).__init__(*args)
        loadUi('pdf_w.ui', self)
        self.btn_file.clicked.connect(self.fileSelect)
        self.btn_split.clicked.connect(self.pdfSplit)
        self.btn_file_merge.clicked.connect(self.fileSelectMerge)
        self.btn_merge.clicked.connect(self.pdfMerge)

    def fileSelect(self):
        file_path = QFileDialog.getOpenFileName(self, "选择文件", "/", "PDF Files (*.pdf)") 
        # print(file_path[0])
        self.filepath.setText(file_path[0])


    def pdfSplit(self):
        read_path = self.filepath.text()
        if read_path:
            file_prefix = os.path.split(read_path)[0]
            try:
                self.textBrowser_log.clear()
                pdf_file = open(read_path, 'rb')
                pdf_input = PdfFileReader(pdf_file)  # 将要分割的PDF内容格式化
                page_count = pdf_input.getNumPages()  # 获取PDF页数
                self.textBrowser_log.append('总页数是' + str(page_count))

                split_std = self.text_split_std.text()
                split_std_list = split_std.split()

                for page_range in split_std_list:
                    # print(page_range)

                    start_page, end_page = page_range.split('-')
                    sp, ep = int(start_page) - 1, int(end_page)
                    if sp < 0:
                    	self.textBrowser_log.append('页数最小从 1 开始，请检查！')
                    if ep > page_count:
                    	self.textBrowser_log.append('页数超过文件限制，请检查！')
                    	break
                    pdf_file_name = os.path.join(file_prefix ,start_page + '-' + end_page + '.pdf')
                    # print(sp, ep, pdf_file_name)
                    
                    try:
                        self.textBrowser_log.append(f'开始分割{sp}页-{ep}页')
                        pdf_output = PdfFileWriter()  # 实例一个 PDF文件编写器
                        for i in range(sp, ep):
                            pdf_output.addPage(pdf_input.getPage(i))
                        with open(pdf_file_name, 'wb') as sub_fp:
                            pdf_output.write(sub_fp)
                        self.textBrowser_log.append(f'完成分割{sp}页-{ep}页，保存为{pdf_file_name}!\n')
                        
                    except IndexError:
                        self.textBrowser_log.append(f'分割页数超过了PDF的页数')
                    finally:
                        sub_fp.close()
                QMessageBox.information(self, '结果', 'PDF分割完毕！')

            except Exception as e:
                self.textBrowser_log.append(e)
            finally:
                pdf_file.close()


    def fileSelectMerge(self):
        pdffiles, suf = QFileDialog.getOpenFileNames(self, "多文件选择", "/", "PDF Files (*.pdf)")
        # print(pdffiles, type(pdffiles), pdffiles[0])
        mutil_file_name = ""
        for x in pdffiles:
            mutil_file_name = mutil_file_name + x + ';'
        self.filepath_merge.setText(mutil_file_name)

 
    def pdfMerge(self):
        mutil_file_paths = self.filepath_merge.text()
        if mutil_file_paths:
            # print(mutil_file_paths)
            mutil_file_path = mutil_file_paths.split(';')
            mutil_file_path = [x for x in mutil_file_path if x != '']
            # print(mutil_file_path)
            output = PdfFileWriter()
            file_prefix = os.path.split(mutil_file_path[0])[0]
            # print(file_prefix)

            self.textBrowser_log.clear()
            

            for pdf_files in mutil_file_path:
                # print(pdf_files)
                # 读取源PDF文件

                pdf_file = open(pdf_files, "rb")
                input = PdfFileReader(pdf_file)

                # 获得源PDF文件中页面总数
                pageCount = input.getNumPages()

                # 分别将page添加到输出output中
                for iPage in range(pageCount):
                    output.addPage(input.getPage(iPage))

                self.textBrowser_log.append(pdf_files + " 合并完成。\n")

            # 写入到目标PDF文件
            outputPath = os.path.join(file_prefix, "合并结果.pdf")
            outputStream = open(outputPath, "wb")
            output.write(outputStream)
            outputStream.close()
            self.textBrowser_log.append("PDF文件合并完成！\n" + outputPath)
            QMessageBox.information(self, '结果', 'PDF合并完毕！')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())