package main

import (
	"bufio"
	"fmt"
	"log"
	"os"
	"strconv"
	"strings"

	"github.com/joho/godotenv"
	"github.com/xuri/excelize/v2"
)

// Hàm lấy giá trị từ ENV, nếu không có thì trả về giá trị mặc định
func getEnvInt(key string, defaultValue int) int {
	val, exists := os.LookupEnv(key)
	if !exists {
		return defaultValue
	}
	intVal, err := strconv.Atoi(val)
	if err != nil {
		log.Fatalf("Giá trị ENV '%s' không hợp lệ: %s", key, val)
	}
	return intVal
}

// Hàm chuyển Rich Text từ Excel thành HTML (chỉ giữ Bold, Italic, Underline, và thay \n thành <br>)
func richTextToHTML(richText []excelize.RichTextRun) string {
	var result strings.Builder

	for _, rt := range richText {
		text := strings.ReplaceAll(rt.Text, "\n", "<br>") // Thay \n bằng <br>
		font := rt.Font

		if font != nil {
			if font.Bold {
				text = "<b>" + text + "</b>"
			}
			if font.Italic {
				text = "<i>" + text + "</i>"
			}
			if font.Underline == "single" {
				text = "<u>" + text + "</u>"
			}
		}

		result.WriteString(text)
	}

	return result.String()
}

func htmlToRichText(htmlText string) []excelize.RichTextRun {
	// Nếu chuỗi rỗng, trả về một đoạn văn bản trống
	if htmlText == "" {
		return []excelize.RichTextRun{{Text: ""}}
	}

	// Thay thế <br> và <br/> bằng \n
	htmlText = strings.ReplaceAll(htmlText, "<br>", "\n")
	htmlText = strings.ReplaceAll(htmlText, "<br/>", "\n")

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
				if tagName == "b" {
					font.Bold = true
				} else if tagName == "i" {
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

// Hàm lưu nội dung vào file .txt
func saveToTextFile(filename, content string) error {
	file, err := os.Create(filename)
	if err != nil {
		return err
	}
	defer file.Close()

	_, err = file.WriteString(content)
	return err
}

func importFromTextToExcel(txtFileName, excelFileName string) {
	// Mở file TXT để đọc
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

		// Kiểm tra nếu dòng bị rỗng
		if strings.TrimSpace(line) == "" {
			continue
		}

		// Tách tên ô và nội dung
		parts := strings.SplitN(line, ": ", 2)
		if len(parts) != 2 {
			log.Printf("Dòng không hợp lệ: %s", line)
			continue
		}

		cellName := parts[0]
		htmlContent := parts[1]

		// Kiểm tra nếu nội dung trống
		if strings.TrimSpace(htmlContent) == "" {
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

	fmt.Printf("File Excel '%s' đã được tạo từ TXT thành công!\n", excelFileName)
}

func main() {
	// Tải biến môi trường từ file .env
	err := godotenv.Load()
	if err != nil {
		log.Fatal("Lỗi khi tải file .env")
	}

	// Lấy tên file Excel từ ENV
	sourceFile := os.Getenv("EXCEL_FILE")
	if sourceFile == "" {
		log.Fatal("Vui lòng đặt biến môi trường 'EXCEL_FILE' trong file .env!")
	}

	// Tên file Excel mới (tạo bản sao)
	targetFile := "copy_" + sourceFile

	// Tên file TXT đầu ra
	outputTxtFile := "output.txt"

	// Tên file Excel sau khi import lại từ TXT
	importedExcelFile := "imported_" + sourceFile

	// Lấy dòng và cột từ ENV (mặc định: đọc từ hàng 1-10, cột A-E)
	startRow := getEnvInt("START_ROW", 1)
	endRow := getEnvInt("END_ROW", 10)
	startCol := getEnvInt("START_COL", 1)
	endCol := getEnvInt("END_COL", 5)

	// Mở file Excel gốc
	srcExcel, err := excelize.OpenFile(sourceFile)
	if err != nil {
		log.Fatalf("Không thể mở file gốc: %s", err)
	}
	defer srcExcel.Close()

	// Tạo file Excel mới
	destExcel := excelize.NewFile()
	sheetName := srcExcel.GetSheetName(0)
	destExcel.NewSheet(sheetName)

	// Chuỗi tổng hợp nội dung để lưu vào file .txt
	var finalContent strings.Builder

	// Duyệt từng ô để copy nội dung và định dạng
	for rowIndex := startRow; rowIndex <= endRow; rowIndex++ {
		for colIndex := startCol; colIndex <= endCol; colIndex++ {
			cellName, _ := excelize.CoordinatesToCellName(colIndex, rowIndex)

			richText, err := srcExcel.GetCellRichText(sheetName, cellName)
			var cellContent string
			if err == nil && richText != nil && len(richText) > 0 {
				destExcel.SetCellRichText(sheetName, cellName, richText)
				cellContent = richTextToHTML(richText)
			} else {
				value, _ := srcExcel.GetCellValue(sheetName, cellName)
				destExcel.SetCellValue(sheetName, cellName, value)
				cellContent = value
			}

			// Lưu style
			styleID, err := srcExcel.GetCellStyle(sheetName, cellName)
			if err == nil {
				destExcel.SetCellStyle(sheetName, cellName, cellName, styleID)
			}

			// Ghi dữ liệu vào file TXT
			finalContent.WriteString(fmt.Sprintf("%s: %s\n", cellName, cellContent))
		}
	}

	// Lưu file Excel mới
	_ = destExcel.SaveAs(targetFile)

	// Lưu nội dung vào file .txt
	_ = saveToTextFile(outputTxtFile, finalContent.String())

	// Chuyển file TXT thành Excel với Rich Text
	importFromTextToExcel(outputTxtFile, importedExcelFile)
}
