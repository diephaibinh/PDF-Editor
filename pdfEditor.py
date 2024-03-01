import pdfProcessing as prcs
import pdfOptimizer as opt


if __name__ == '__main__':
    print("1: Ghép file đơn lẻ (Single file merge)")
    print("2: Ghép nhiều file (Multiple files merge)")
    print("3: Chèn trang (Insert)")
    print("4: Xóa trang (Remove)")
    print("5: Xoay trang (Rotate)")
    print("6: Trích trang (Extract)")
    print("7: Tối ưu file (Optimze)")
    print("0: Thoát (Exit)")
        
    choice = "0"
    choice = input("Lựa chọn thao tác (Nhập vào 1 số + enter): ")
        
    if choice == '1':
        prcs.execute_SingleMerge()
    if choice == '2':
        prcs.execute_MultipleMerge()
    if choice == '3':
        prcs.execute_Insert()
    if choice == '4':
        prcs.execute_Remove()
    if choice == '5':
        prcs.execute_Rotate()
    if choice == '6':
        prcs.execute_Extract()
    if choice == '7':
        filename = input("Enter txt file name: ")
        opt.main(filename)
