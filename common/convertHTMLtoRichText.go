package common

import (
	"bufio"
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

func htmlToRichText(htmlText string) []excelize.RichTextRun {
	// Nếu chuỗi rỗng, trả về một đoạn văn bản trống
	if htmlText == "" {
		return []excelize.RichTextRun{{Text: ""}}
	}

	// Thay thế <br>, <br/> và <p> bằng \n
	htmlText = strings.ReplaceAll(htmlText, "<br>", "\n")
	htmlText = strings.ReplaceAll(htmlText, "<br/>", "\n")
	// Thay thế <p> thành "" và </p> thành \n
	htmlText = strings.ReplaceAll(htmlText, "<p>", "")
	htmlText = strings.ReplaceAll(htmlText, "</p>", "")

	// Stack lưu trữ các thẻ mở
	type tagInfo struct {
		tag   string
		start int
	}
	var tagStack []tagInfo
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

			// Lấy nội dung của thẻ (VD: "b", "/b")
			tagContent := htmlText[i+1 : i+endTagIndex]
			isClosingTag := strings.HasPrefix(tagContent, "/")
			tagName := strings.TrimPrefix(tagContent, "/")

			// Nếu có văn bản trước thẻ, thêm vào richText
			if textBuffer.Len() > 0 {
				richText = append(richText, excelize.RichTextRun{Text: textBuffer.String()})
				textBuffer.Reset()
			}

			// Nếu là thẻ mở, đẩy vào stack
			if !isClosingTag {
				tagStack = append(tagStack, tagInfo{tag: tagName, start: len(richText)})
			} else {
				// Nếu là thẻ đóng, kiểm tra xem có thẻ mở khớp không
				if len(tagStack) == 0 || tagStack[len(tagStack)-1].tag != tagName {
					log.Printf("Cảnh báo: Thẻ đóng </%s> không có thẻ mở tương ứng!", tagName)
					i += endTagIndex + 1
					continue
				}

				// Lấy phần văn bản bên trong thẻ
				startIndex := tagStack[len(tagStack)-1].start
				tagStack = tagStack[:len(tagStack)-1] // Loại bỏ thẻ đã đóng

				// Áp dụng định dạng cho nội dung trong thẻ
				font := &excelize.Font{}
				if tagName == "b" || tagName == "strong" {
					font.Bold = true
				} else if tagName == "i" || tagName == "em" {
					font.Italic = true
				} else if tagName == "u" {
					font.Underline = "single"
				}

				// Cập nhật danh sách RichTextRun
				for j := startIndex; j < len(richText); j++ {
					if richText[j].Font == nil {
						richText[j].Font = &excelize.Font{}
					}
					if font.Bold {
						richText[j].Font.Bold = true
					}
					if font.Italic {
						richText[j].Font.Italic = true
					}
					if font.Underline == "single" {
						richText[j].Font.Underline = "single"
					}
				}
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
func ImportFromTextToExcel(txtFileName, excelFileName string) {
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

// func main() {
// 	// Đọc từ TXT và ghi vào Excel
// 	ImportFromTextToExcel("output.txt", "imported.xlsx")
// }
