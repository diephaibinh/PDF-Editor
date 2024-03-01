from PIL import Image, ImageFilter
import shutil
import fitz
import os
import io
import time

# maximum 9999 pages
# just for only image pdf


DEFAULT_QUALITY = 50
DEFAULT_DPI = (300, 300)
DEFAULT_SIZE_RATIO = 100
DEFAULT_IMAGE_RADIUS = 0
DEFAULT_IMAGE_FILTER_SIZE = 1
DEFAULT_IMAGE_DIR = "image_page"


# ------------------------------------------- Extract Image -------------------------------------------
# folder is input folder name. (just name)
def extract_img(pdf_filename, folder):
    try:
        shutil.rmtree(folder)
        os.mkdir(folder)
    except Exception:
        os.mkdir(folder)

    doc = fitz.open(pdf_filename)
    is_jpg = True

    for page_idx, page in enumerate(doc, start=1):
        try:
            images = page.get_images()
            for img_idx, img in enumerate(images, start=1):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                img_ext = base_image["ext"]
                if str(img_ext).lower() != 'jpg' or str(img_ext).lower() != 'jpeg':
                    is_jpg = False

                image = Image.open(io.BytesIO(image_bytes))
                filename = f"{page_idx:04d}_{img_idx:04d}.{str(img_ext)}"
                image.save(open(folder + "\\" + filename, "wb"))
        except Exception:
            continue

    return is_jpg


# ------------------------------------------- Optimize Image -------------------------------------------
def optimize_img(input_img, output_img, quality, dpi, img_size_ratio, radius, filter_size):
    img = Image.open(input_img)
    w = int(img.size[0] * img_size_ratio / 100)
    h = int(img.size[1] * img_size_ratio / 100)
    img = img.resize((w, h), Image.LANCZOS)

    img_output = img.filter(ImageFilter.GaussianBlur(radius=radius))
    img_output = img_output.filter(ImageFilter.MinFilter(size=filter_size))
    img_output.save(output_img, optimize=True, quality=quality, dpi=dpi)


def optimize_img_dir(input_dir, output_dir, quality, dpi, img_size_ratio, radius, filter_size):
    img_lst = os.listdir(input_dir)
    for image_name in img_lst:
        optimize_img(
            f"{input_dir}\\{image_name}",
            f"{output_dir}\\{image_name}",
            quality, dpi, img_size_ratio, radius, filter_size
        )


# ------------------------------------------- Convert to JPG -------------------------------------------
def convert_to_jpg(input_img, output_img):
    img = Image.open(input_img)
    mode = img.mode
    if mode[len(mode) - 1] == 'A':
        img = img.convert(mode[:len(mode) - 1])
        img.save(output_img)

    new_img = os.path.splitext(output_img)[0] + "-temp.jpeg"
    img.save(new_img, format="JPEG")

    os.remove(input_img)
    os.rename(new_img, new_img.replace("-temp", ""))


def convert_to_jpg_dir(input_dir, output_dir):
    img_lst = os.listdir(input_dir)
    for image_name in img_lst:
        convert_to_jpg(f"{input_dir}\\{image_name}", f"{output_dir}\\{image_name}")


# ------------------------------------------- Replace Image -------------------------------------------
def update_image_from_dir(pdf_filename, new_filename, image_dir):
    img_list = os.listdir(image_dir)
    doc = fitz.open(pdf_filename)

    for page_idx, page in enumerate(doc, start=1):
        images = page.get_images(full=True)
        opt_imgs = [img for img in img_list if img.startswith(f"{page_idx:04d}")]
        opt_imgs.sort()
        for img_idx, img in enumerate(images):
            rect = page.get_image_bbox(img)
            image_name = image_dir + "\\" + opt_imgs[img_idx]
            page.insert_image(rect, filename=image_name)

        page.clean_contents()
        xref = page.get_contents()[0]
        cont_lines = doc.xref_stream(xref).splitlines()
        for j in range(len(cont_lines)):
            line = cont_lines[j]
            if line.startswith(b"/Im") and line.endswith(b" Do"):
                cont_lines[j] = b""
        cont = b"\n".join(cont_lines)
        doc.update_stream(xref, cont)
        page.clean_contents()

    doc.save(new_filename, garbage=3, deflate=True)


# ------------------------------------------- Final -------------------------------------------
def compress_pdf(input_pdf, output_pdf, quality, dpi, img_size_ratio, radius, filter_size, img_dir=DEFAULT_IMAGE_DIR):
    print("\nFile:", os.path.basename(input_pdf))

    try:
        if not os.path.exists(os.path.dirname(output_pdf)):
            os.mkdir(os.path.dirname(output_pdf))

        start = time.time()
        img = extract_img(input_pdf, img_dir)
        if not img:
            convert_to_jpg_dir(img_dir, img_dir)
        optimize_img_dir(img_dir, img_dir, quality, dpi, img_size_ratio, radius, filter_size)
        update_image_from_dir(input_pdf, output_pdf, img_dir)
        end = time.time()

        # Info stat
        before = round(os.path.getsize(input_pdf) / (1024 ** 2), 2)
        after = round(os.path.getsize(output_pdf) / (1024 ** 2), 2)
        ratio = round(100 * (1 - after / before), 2)
        print(before, "MB ---->", after, "MB")
        print("Ratio:", ratio, "%")
        print("Time:", round(end - start, 1))

        shutil.rmtree(img_dir)
    except Exception:
        print("Having some problem")


def compress_pdf_dir(input_dir, output_dir, quality, dpi, img_size_ratio, radius, filter_size,
                     img_dir=DEFAULT_IMAGE_DIR):
    if not os.path.exists(output_dir):
        os.mkdir(output_dir)
    file = os.listdir(input_dir)
    for f in file:
        input_pdf = input_dir + "\\" + f
        if os.path.splitext(input_pdf)[1] == ".pdf":
            output_pdf = output_dir + "\\" + f
            compress_pdf(input_pdf, output_pdf, quality, dpi, img_size_ratio, radius, filter_size, img_dir)
        else:
            continue


# ------------------------------------------- Read information from txt file -------------------------------------------
######################## Choice ########################
# "/" = 47, "\" = 92
# 1: compress file
# 2: compress dir
def read_file(filename):
    elements = open(filename, "r", encoding="utf-8").read().strip().split("\n\n")
    lst = []
    for e in elements:
        dic = {
            "choice": 1,
            "input": None,
            "output": None,
            "quality": DEFAULT_QUALITY,
            "dpi": DEFAULT_DPI,
            "image_size": DEFAULT_SIZE_RATIO,
            "radius": DEFAULT_IMAGE_RADIUS,
            "filter_size": DEFAULT_IMAGE_FILTER_SIZE,
        }

        lines = e.split("\n")
        start = 2
        dic["input"] = lines[0]
        output = None

        if os.path.isdir(dic["input"]):
            dic["choice"] = 2

        # Find output path
        if dic["choice"] == 1:
            output = os.path.dirname(dic["input"]) + "\\compressed"
        if dic["choice"] == 2:
            output = dic["input"] + "\\compressed"
        if len(lines) > 1:  # May have output path
            if not os.path.isdir(lines[1]):
                dic["output"] = output
                start = 1
            else:
                dic["output"] = lines[1]
        else:  # Just have input path, no more info
            dic["output"] = output
            start = 1

        for i in range(start, len(lines)):
            parameter = lines[i].strip().split('\t')
            if parameter[0].lower() == "dpi":
                dic[parameter[0].lower()] = eval(parameter[1])
            else:
                dic[parameter[0].lower()] = int(parameter[1])

        lst.append(dic)

    return lst


def main(file):
    elements = read_file(file)
    for dic in elements:
        if dic["choice"] == 1:
            output_pdf = dic["output"] + "\\" + os.path.basename(dic["input"])
            compress_pdf(
                dic["input"],
                output_pdf,
                dic["quality"],
                dic["dpi"],
                dic["image_size"],
                dic['radius'],
                dic["filter_size"],
            )
        if dic["choice"] == 2:
            compress_pdf_dir(
                dic["input"],
                dic["output"],
                dic["quality"],
                dic["dpi"],
                dic["image_size"],
                dic['radius'],
                dic["filter_size"]
            )


if __name__ == '__main__':
    filename = input("Enter txt file name (Enter '0' to exit): ")

    if filename != '0':
        start = time.time()
        main(filename)
        end = time.time()

        try:
            shutil.rmtree(DEFAULT_IMAGE_DIR)
        except:
            ""

        print("\nDone!!")
        print("Time:", round(end - start, 2), 's')

# optimize_img_dir("image_page", "image_page1", 10, (300, 300), 100, 0, 3)
# compress_pdf("test.pdf", "D:\\Tools\\PDF Editor\\test_opt.pdf", DEFAULT_QUALITY, DEFAULT_DPI, DEFAULT_SIZE_RATIO, DEFAULT_IMAGE_RADIUS, DEFAULT_IMAGE_FILTER_SIZE)
