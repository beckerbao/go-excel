package main

import (
	"bufio"
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Hàm chuyển đổi HTML sang Rich Text của Excel
func htmlToRichText(htmlText string) []excelize.RichTextRun {
	var richText []excelize.RichTextRun
	var textBuffer strings.Builder
	i := 0

	for i < len(htmlText) {
		// Nếu gặp thẻ mở
		if htmlText[i] == '<' {
			endTagIndex := strings.Index(htmlText[i:], ">")
			if endTagIndex == -1 {
				// Không tìm thấy dấu ">", xử lý như văn bản bình thường
				textBuffer.WriteByte(htmlText[i])
				i++
				continue
			}

			// Lấy nội dung của thẻ
			tagContent := htmlText[i+1 : i+endTagIndex]
			tagName := strings.TrimPrefix(tagContent, "/")

			// Nếu có văn bản trước thẻ, thêm vào richText
			if textBuffer.Len() > 0 {
				richText = append(richText, excelize.RichTextRun{Text: textBuffer.String()})
				textBuffer.Reset()
			}

			// Nếu là thẻ mở, xử lý font
			font := &excelize.Font{}
			if tagName == "b" {
				font.Bold = true
			} else if tagName == "i" {
				font.Italic = true
			} else if tagName == "u" {
				font.Underline = "single"
			}

			// Xử lý nội dung bên trong thẻ
			nestedText := htmlToRichText(htmlText[i+endTagIndex+1:])

			// Áp dụng định dạng lên tất cả phần tử con
			for j := range nestedText {
				if nestedText[j].Font == nil {
					nestedText[j].Font = &excelize.Font{}
				}
				if font.Bold {
					nestedText[j].Font.Bold = true
				}
				if font.Italic {
					nestedText[j].Font.Italic = true
				}
				if font.Underline == "single" {
					nestedText[j].Font.Underline = "single"
				}
				richText = append(richText, nestedText[j])
			}

			// Cập nhật vị trí duyệt
			i += endTagIndex + 1
		} else {
			// Nếu là văn bản thường, thêm vào buffer
			textBuffer.WriteByte(htmlText[i])
			i++
		}
	}

	// Thêm phần còn lại của văn bản nếu có
	if textBuffer.Len() > 0 {
		richText = append(richText, excelize.RichTextRun{Text: textBuffer.String()})
	}

	return richText
}

// Hàm import từ TXT vào Excel (BỎ QUA HÌNH ẢNH)
func importFromTextToExcel(txtFileName, excelFileName string) {
	// Mở file TXT
	file, err := os.Open(txtFileName)
	if err != nil {
		log.Fatalf("Không thể mở file TXT: %s", err)
	}
	defer file.Close()

	// Tạo file Excel mới
	excelFile := excelize.NewFile()
	sheetName := "Sheet1"
	excelFile.NewSheet(sheetName)

	// Đọc từng dòng trong file TXT
	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		line := scanner.Text()

		// Kiểm tra nếu dòng rỗng
		if strings.TrimSpace(line) == "" {
			continue
		}

		// Tách ô và nội dung (VD: "A2: Nội dung")
		parts := strings.SplitN(line, ": ", 2)
		if len(parts) != 2 {
			log.Printf("Dòng không hợp lệ: %s", line)
			continue
		}

		cellName := parts[0]
		htmlContent := parts[1]

		// Bỏ qua nội dung có chứa hình ảnh `<img>` (VÌ ĐÃ LOẠI BỎ)
		if strings.Contains(htmlContent, "<img") {
			continue
		}

		// Chuyển đổi HTML sang Rich Text
		richText := htmlToRichText(htmlContent)

		// Ghi dữ liệu vào file Excel
		err := excelFile.SetCellRichText(sheetName, cellName, richText)
		if err != nil {
			log.Printf("Lỗi khi ghi vào %s: %s", cellName, err)
		}
	}

	// Lưu file Excel
	if err := excelFile.SaveAs(excelFileName); err != nil {
		log.Fatalf("Không thể lưu file Excel: %s", err)
	}

	fmt.Printf("File Excel '%s' đã được tạo thành công!\n", excelFileName)
}

func main() {
	// Đọc từ TXT và ghi vào Excel
	importFromTextToExcel("output.txt", "imported.xlsx")
}
