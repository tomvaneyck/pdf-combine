import sys
from os import startfile
from pathlib import Path
from typing import Tuple

import easygui
from pypdf import PageObject, PaperSize, PdfReader, PdfWriter, Transformation


def determine_scale(src_w, src_h, dest_w, dest_h, fit):
    if not fit:
        if src_w - dest_w > 0.1 * dest_w or src_h - dest_h > 0.1 * dest_h:
            return min(dest_h / src_h, dest_w / src_w)
        else:
            return 1
    else:
        scale_factor = max(dest_h / src_h, dest_w / src_w)
        if src_w * scale_factor - dest_w > 0.1 * dest_w or src_h * scale_factor - dest_h > 0.1 * dest_h:
            return min(dest_h / src_h, dest_w / src_w)
        else:
            return scale_factor


def merge_slides_page(
    srcpage: PageObject, index, pagenum, destpage: PageObject | None, fit, writer: PdfWriter
) -> Tuple[int, PageObject]:
    convmm = 0.352777
    hmar = 5
    vmar = 25
    sw = 99
    sh = 52
    hpad = 2
    vpad = 45
    w = 210
    h = 297

    #
    # +---------------------------------------------------------------------+
    # |    ʌ                                                                |
    # |    25 mm                                                            |
    # |    v <         99mm          >< 2mm >                               |
    # |<5mm>+------------------------+       +------------------------+     |
    # |     | ʌ                      |       |                        |     |
    # |     |                        |       |                        |     |
    # |     | 52 mm                  |       |                        |     |
    # |     |                        |       |                        |     |
    # |     |                        |       |                        |     |
    # |     | v                      |       |                        |     |
    # |     +------------------------+       +------------------------+     |
    # |     ʌ                                                               |
    # |     45mm                                                            |
    # |                                                                     |
    # |     v                                                               |
    # |     +------------------------+                                      |
    # |     | ʌ                      |                                      |
    # |     |                        |                                      |
    # |     | 52 mm                  |                                      |
    # |     |                        |                                      |
    # |     |                        |                                      |
    # |     | v                      |                                      |
    # |     +------------------------+                                      |
    # |                                                                     |
    # |                                                                     |
    # |                                                                     |
    # |                                                                     |
    # |                                                                     |
    # |     +------------------------+                                      |
    # |     | ʌ                      |                                      |
    # |     |                        |                                      |
    # |     | 52 mm                  |                                      |
    # |     |                        |                                      |
    # |     |                        |                                      |
    # |     | v                      |                                      |
    # |     +------------------------+                                      |
    # |                                                                     |
    # |                                                                     |
    # |                                                                     |
    # +---------------------------------------------------------------------+
    #

    if index % 6 == 0:
        destpage = writer.add_blank_page(
            width=PaperSize.A4.width, height=PaperSize.A4.height
        )
        pagenum += 1

    if destpage is None:
        sys.exit("No destination page")

    column = index % 2
    line = (index % 6) // 2
    xpos = (column * (sw + hpad) + hmar) / convmm
    ypos = (h - vmar - line * (sh + vpad) - sh) / convmm

    destpage.merge_transformed_page(
        srcpage,
        Transformation()
        .scale(sw / (srcpage.mediabox.width * convmm))
        .translate(xpos, ypos),
    )

    return (pagenum, destpage)


def merge_single_page(
    srcpage: PageObject, _, pagenum, destpage: PageObject | None, fit, writer: PdfWriter
) -> Tuple[int, PageObject]:
    destpage = writer.add_blank_page(
        width=PaperSize.A4.width, height=PaperSize.A4.height
    )
    pagenum += 1

    A4_w = PaperSize.A4.width
    A4_h = PaperSize.A4.height

    # resize page to fit *inside* A4
    w = srcpage.mediabox.width
    h = srcpage.mediabox.height

    scale_factor = determine_scale(w, h, A4_w, A4_h, fit)
    trans_w = (A4_w - scale_factor * w) / 2
    trans_h = (A4_h - scale_factor * h) / 2

    transform = (
        Transformation().scale(scale_factor, scale_factor).translate(trans_w, trans_h)
    )
    destpage.merge_transformed_page(srcpage, transform)

    return (pagenum, destpage)


def merge_double_page(
    srcpage: PageObject, index, pagenum, destpage: PageObject | None, fit, writer: PdfWriter
) -> Tuple[int, PageObject]:
    if index % 2 == 0:
        destpage = writer.add_blank_page(
            width=PaperSize.A4.height, height=PaperSize.A4.width
        )
        pagenum += 1

    if destpage is None:
        sys.exit("No destination page")

    column = index % 2
    xpos = column * destpage.mediabox.width / 2

    final_w = PaperSize.A4.height / 2
    final_h = PaperSize.A4.width
    w = srcpage.mediabox.width
    h = srcpage.mediabox.height

    # resize page to fit *inside* half a A4
    scale_factor = determine_scale(w, h, final_w, final_h, fit)
    # scale_factor = min(final_h / h, final_w / w)
    trans_w = (final_w - scale_factor * w) / 2
    trans_h = (final_h - scale_factor * h) / 2

    transform = (
        Transformation()
        .scale(scale_factor, scale_factor)
        .translate(trans_w + xpos, trans_h)
    )
    destpage.merge_transformed_page(srcpage, transform)

    return (pagenum, destpage)


def create_pdf(pdfs, format, size, writer: PdfWriter):
    pagenum = 1
    destpage = None
    for pdf in pdfs:
        print("Merging pdf " + pdf)
        if pagenum % 2 == 0:
            if format == 'double':
                writer.add_blank_page(width=PaperSize.A4.height, height=PaperSize.A4.width)
            else:
                writer.add_blank_page(width=PaperSize.A4.width, height=PaperSize.A4.height)
            pagenum += 1

        # Add bookmark
        path = Path(pdf)
        writer.add_outline_item(path.name, page_number=pagenum + 1)

        for i, source in enumerate(PdfReader(pdf, strict=True).pages):
            if format == "slides":
                (pagenum, destpage) = merge_slides_page(
                    source, i, pagenum, destpage, size, writer
                )
            elif format == "single":
                (pagenum, destpage) = merge_single_page(
                    source, i, pagenum, destpage, size, writer
                )
            elif format == "double":
                (pagenum, destpage) = merge_double_page(
                    source, i, pagenum, destpage, size, writer
                )

    for i, page in enumerate(writer.pages):
        print("Compressing page " + str(i + 1))
        page.compress_content_streams()


if __name__ == "__main__":
    output = PdfWriter()

    pdfs = easygui.fileopenbox(
        multiple=True, title="Select pdfs to combine", default="*.pdf"
    )
    if pdfs is None:
        sys.exit("No input files selected")
    output = easygui.filesavebox(title="Select output file", default="*.pdf")
    if output is None or type(output) is list:
        sys.exit("No output file selected")
    elif output[-4:] != ".pdf":
        output += ".pdf"

    formats = ["single", "double", "slides"]
    format = easygui.choicebox(msg="Select the final format", choices=formats)
    if format not in formats:
        sys.exit("No output format selected")

    sizes = ["fit", "real"]
    size = easygui.choicebox(msg="Select the final fit of each page", choices=sizes)
    if size not in sizes:
        sys.exit("No output size selected")

    writer = PdfWriter()
    pdf = create_pdf(pdfs, format, size == "fit", writer)

    with open(output, "wb") as pdf:
        writer.write(pdf)

    startfile(output)
