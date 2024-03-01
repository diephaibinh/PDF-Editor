from PyPDF3 import PdfFileWriter, PdfFileReader
import os



# ---------------------------------1---------------------------------
def mergebyPage(NumOfFile=None, outputFile=None):
    try:
        NumOfFile = int(input("Number of file: "))
        if NumOfFile == 0:
            return False
        outputFile = input("New file name: ")

        lstPDF = []
        count = 1

        while count <= NumOfFile:
            filename = input("File " + str(count) + ": ")
            begin = int(input("Begin: "))
            end = int(input("End: "))
            if (filename == "0"):
                return False
            lstPDF.append((filename, begin, end))
            count += 1
        print(lstPDF)

        output = PdfFileWriter()
        for pdf in lstPDF:
            inputPDF = PdfFileReader(open(pdf[0], "rb"))
            for page in range(pdf[1] - 1, pdf[2]):
                output.addPage(inputPDF.getPage(page))
        outputPDF = open(outputFile, "wb")
        output.write(outputPDF)
        outputPDF.close()
        return True
    except:
        print("Lỗi! Kiểm tra lại thông tin đầu vào.\n")
        return False


# ---------------------------------2---------------------------------
def mergeSingleFile(lst):
    outputPDF = open(lst[0], "wb")
    output = PdfFileWriter()

    for i in range(1, len(lst)):
        inputPDF = PdfFileReader(lst[i][0])
        for page in range(int(lst[i][1]) - 1, int(lst[i][2])):
            output.addPage(inputPDF.getPage(page))

    output.write(outputPDF)
    outputPDF.close()


def merge_bytxtfile(txtFile):
    pdf = open(txtFile, "r", encoding='utf-8').read().strip().split("\n\n")
    lst = []

    for file in pdf:
        info = file.split("\n")
        info[0] = info[0].strip()
        for i in range(1, len(info)):
            info[i] = tuple(info[i].strip().split("\t"))
        lst.append(info)

    for file in lst:
        print("Đang thực hiện file", os.path.basename(file[0]))
        try:
            mergeSingleFile(file)
            print("Hoàn thành.\n")
        except:
            print("Lỗi! Kiểm tra lại thông tin đầu vào.\n")



# ---------------------------------3---------------------------------
def insert(pdf_file1, index, pdf_file2, page):
    try:
        newfile = os.path.splitext(pdf_file1)[0] + "_NewFile.pdf"
        inputPDF1 = PdfFileReader(pdf_file1)
        inputPDF2 = PdfFileReader(pdf_file2)

    # if (index - 1) < 0 or (index - 1) > inputPDF1.getNumPages() or (page - 1) < 0 or (page - 1) >= inputPDF2.getNumPages():
    #     return False

        output = PdfFileWriter()

        if index - 1 == inputPDF1.getNumPages():
            for pagenum in range(inputPDF1.getNumPages()):
                output.addPage(inputPDF1.getPage(pagenum))
            output.addPage(inputPDF2.getPage(page - 1))
        else:
            pagenum = 0
            for pagenum in range(inputPDF1.getNumPages()):
                if pagenum == index - 1:
                    output.addPage(inputPDF2.getPage(page - 1))
                    break
                output.addPage(inputPDF1.getPage(pagenum))
            for i in range(pagenum, inputPDF1.getNumPages()):
                output.addPage(inputPDF1.getPage(i))

        outputPDF = open(newfile, "wb")
        output.write(outputPDF)
        outputPDF.close()

        os.remove(pdf_file1)
        os.rename(newfile, pdf_file1)
        return True
    except:
        print("Lỗi! Kiểm tra lại thông tin đầu vào.\n")
        return False


# ---------------------------------4---------------------------------
def remove(pdf_file, page):
    try:
        newfile = os.path.splitext(pdf_file)[0] + "_NewFile.pdf"
        inputPDF = PdfFileReader(pdf_file)
        output = PdfFileWriter()
            
        pagenum = 0
        for pagenum in range(inputPDF.getNumPages()):
            if pagenum == page - 1:
                pagenum += 1
                break
            output.addPage(inputPDF.getPage(pagenum))
        for i in range(pagenum, inputPDF.getNumPages()):
            output.addPage(inputPDF.getPage(i))

        outputPDF = open(newfile, "wb")
        output.write(outputPDF)
        outputPDF.close()
        return True
    except:
        print("Lỗi! Kiểm tra lại thông tin đầu vào.\n")
        return False


# ---------------------------------5---------------------------------
# choice: a (right), b (left), c (180)
def rotate(pdf_file, page_rotate, choice):
    try:
        newfile = os.path.splitext(pdf_file)[0] + "_NewFile.pdf"
        inputPDF = PdfFileReader(pdf_file)

        if (page_rotate - 1) >= 0 and (page_rotate - 1) < inputPDF.getNumPages():
            output = PdfFileWriter()
            pagenum = 0
            for pagenum in range(inputPDF.getNumPages()):
                page = inputPDF.getPage(pagenum)
                if pagenum == page_rotate - 1:
                    if choice == 'a' or choice == 'b' or choice == 'c':
                        if choice == "a":
                            page.rotateClockwise(90)
                        if choice == "b":
                            page.rotateClockwise(270)
                        if choice == "c":
                            page.rotateClockwise(180)
                        output.addPage(page)
                        pagenum += 1
                        break
                    else:
                        print("Lỗi! Kiểm tra lại thông tin đầu vào.\n")
                        return False
                output.addPage(page)

            for i in range(pagenum, inputPDF.getNumPages()):
                output.addPage(inputPDF.getPage(i))

            outputPDF = open(newfile, "wb")
            output.write(outputPDF)
            outputPDF.close()

            os.remove(pdf_file)
            os.rename(newfile, pdf_file)
        return True
    except:
        print("Lỗi! Kiểm tra lại thông tin đầu vào.\n")
        return False


# ---------------------------------6---------------------------------
def extract(pdf_file, page):
    try:
        newfile = os.path.splitext(pdf_file)[0] + "_page" + str(page) + ".pdf"
        inputPDF = PdfFileReader(pdf_file)

        output = PdfFileWriter()
        output.addPage(inputPDF.getPage(page - 1))

        outputPDF = open(newfile, "wb")
        output.write(outputPDF)
        outputPDF.close()
        return True
    except:
        print("Lỗi! Kiểm tra lại thông tin đầu vào.\n")
        return False


# ---------------------------------Execute---------------------------------
def execute_SingleMerge():
    mergebyPage()

def execute_MultipleMerge():
    filename = input("Nhập tên file dữ liệu (.txt file): ")
    merge_bytxtfile(filename)

def execute_Insert():
    pdf_file1 = input("Nhập đường dẫn file cần chèn vào: ")
    index = int(input("Nhập vị trí cần chèn: "))
    pdf_file2 = input("Nhập đường dẫn file để lấy trang chèn: ")
    page = int(input("Nhập số trang (trang 1, trang 2,...--> chỉ nhập số): "))
    insert(pdf_file1, index, pdf_file2, page)

def execute_Remove():
    pdf_file = input("Nhập đường dẫn file: ")
    page = int(input("Nhập số trang: "))
    remove(pdf_file, page)

def execute_Rotate():
    pdf_file = input("Nhập đường dẫn file: ")
    page = int(input("Nhập số trang: "))
    print("Nhập lựa chọn:")
    print("   a) Xoay phải 90 độ")
    print("   b) Xoay trái 90 độ")
    print("   c) Xoay 180 độ")
    print("   d) Thoát")
    rotate_mode = input()
    rotate(pdf_file, page, rotate_mode)

def execute_Extract():
    pdf_file = input("Nhập đường dẫn file: ")
    page = int(input("Nhập số trang cần trích (trang 1, trang 2,...--> chỉ nhập số): "))
    extract(pdf_file, page)




if __name__ == '__main__':
    print("1: Ghép file đơn lẻ (Single file merge)")
    print("2: Ghép nhiều file (Multiple files merge)")
    print("3: Chèn trang (Insert)")
    print("4: Xóa trang (Remove)")
    print("5: Xoay trang (Rotate)")
    print("6: Trích trang (Extract)")
    print("0: Thoát (Exit)")
        
    choice = 0
    choice = input("Lựa chọn thao tác (Nhập vào 1 số + enter): ")
        
    if choice == '1':
        execute_SingleMerge()
    if choice == '2':
        execute_MultipleMerge()
    if choice == '3':
        execute_Insert()
    if choice == '4':
        execute_Remove()
    if choice == '5':
        execute_Rotate()
    if choice == '6':
        execute_Extract()
