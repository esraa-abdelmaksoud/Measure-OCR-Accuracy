import cv2
import os
import xlsxwriter

rect_folder_path = r'/mnt/D/Upwork/CompleteMap-Rom/OCR_esraa/red_selected/'
org_folder_path = r'/mnt/D/Upwork/CompleteMap-Rom/OCR_esraa/image_selected/'
output_folder_path = r'/mnt/D/Upwork/CompleteMap-Rom/OCR_esraa/excel_rect/'
conv_folder_path = r'/mnt/D/Upwork/CompleteMap-Rom/OCR_esraa/color_conversion/'
snips_folder_path = r'/mnt/D/Upwork/CompleteMap-Rom/OCR_esraa/red_rect_snips/'
# rect_folder_path = r'/mnt/D/Upwork/CompleteMap-Rom/OCR_esraa/test/red_selected/'
# org_folder_path = r'/mnt/D/Upwork/CompleteMap-Rom/OCR_esraa/test/image_selected/'
# output_folder_path = r'/mnt/D/Upwork/CompleteMap-Rom/OCR_esraa/test/rect_output/'
# conv_folder_path = r'/mnt/D/Upwork/CompleteMap-Rom/OCR_esraa/test/color_conversion/'
# snips_folder_path = r'/mnt/D/Upwork/CompleteMap-Rom/OCR_esraa/test/red_rect_snips/'


def main(output_folder_path, rect_folder_path):
    files = os.listdir(rect_folder_path)
    for file in files:
            detect_rects(file,rect_folder_path, org_folder_path, output_folder_path)

def detect_rects(file,rect_folder_path, org_folder_path, output_folder_path):
    rect_img_path = os.path.join(rect_folder_path,file)
    org_img_path = os.path.join(org_folder_path, file)
    excel_path = os.path.join(output_folder_path, f'{file[:-4]}.xlsx')
    workbook = xlsxwriter.Workbook(excel_path)
    worksheet = workbook.add_worksheet()
    cols = [
            "file",
            "page",
            "block",
            "x0",
            "y0",
            "x1",
            "y1",
            "text",
            "transcription",
            "image",
        ]
    for c in range(len(cols)):
        worksheet.write(0, c, cols[c])

    # Read images
    rect_img = cv2.imread(rect_img_path)
    org_img = cv2.imread(org_img_path)
    # print(rect_img_path)
    # print(org_img_path)

    green = rect_img[:,:,1]
    green_img_path = os.path.join(conv_folder_path, f'{file[:-4]}-green.png')
    cv2.imwrite(green_img_path,green)
    _,thresh = cv2.threshold(green,1,255,cv2.THRESH_BINARY,)
    thresh_img_path = os.path.join(conv_folder_path, f'{file[:-4]}-thresh.png')
    cv2.imwrite(thresh_img_path,thresh)

    contours,_ = cv2.findContours(thresh, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    print(file,'Contours:', len(contours))
    for c in range(len(contours)-1):
        # print(c)
        x,y,w,h = cv2.boundingRect(contours[c])
        snip = org_img[y:y+h,x:x+w]
        if snip.shape[0] > 10 and snip.shape[1] > 10:
            snip_path = os.path.join(snips_folder_path, f'{file[:-4]}-snip-{c}.png')
            cv2.imwrite(snip_path,snip)
            # print('snip path:', snip_path)

            # Adjusting worksheet
            worksheet.set_row(c+1, ((y+h) - y) // 2.5)

            # Writing data to worksheet
            worksheet.insert_image(
                f"J{c+2}", snip_path, {"x_scale": 0.5, "y_scale": 0.5}
            )
            page_idx = file.rfind("-")
            fname = file[:page_idx]
            page_num = file[page_idx+1:]

            values = [fname, page_num, "UNK", x, y, x+w, y+h, ""]
            for v in range(len(values)):
                worksheet.write(c+1, v, values[v])
    workbook.close()
    thresh = cv2.cvtColor(thresh,cv2.COLOR_GRAY2BGR)
    cv2.drawContours(thresh, contours, -1, (255, 0, 0), 3)
    cont_img_path = os.path.join(conv_folder_path, f'{file[:-4]}-contour.png')
    cv2.imwrite(cont_img_path,thresh)



if __name__ == "__main__":
    main(output_folder_path, rect_folder_path)