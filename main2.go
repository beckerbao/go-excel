package main

import (
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/joho/godotenv"
	"github.com/xuri/excelize/v2"
)

// Hàm chuyển Rich Text từ Excel thành HTML (chỉ giữ Bold, Italic, Underline)
func richTextToHTML(richText []excelize.RichTextRun) string {
	var result strings.Builder

	for _, rt := range richText {
		text := rt.Text
		font := rt.Font

		if font != nil {
			// Nếu Bold == true, bọc trong <b>
			if font.Bold {
				text = "<b>" + text + "</b>"
			}
			// Nếu Italic == true, bọc trong <i>
			if font.Italic {
				text = "<i>" + text + "</i>"
			}
			// Nếu Underline có giá trị, bọc trong <u>
			if font.Underline == "single" {
				text = "<u>" + text + "</u>"
			}
		}

		result.WriteString(text)
	}

	return result.String()
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

func main() {
	// Tải biến môi trường từ file .env
	err := godotenv.Load()
	if err != nil {
		log.Fatal("Lỗi khi tải file .env")
	}

	// Lấy tên file Excel từ biến môi trường
	filePath := os.Getenv("EXCEL_FILE")
	if filePath == "" {
		log.Fatal("Vui lòng đặt biến môi trường 'EXCEL_FILE' trong file .env!")
	}

	// Tên file TXT đầu ra
	outputFile := "output.txt"

	// Mở file Excel
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		log.Fatalf("Không thể mở file Excel: %s", err)
	}
	defer f.Close()

	// Lấy tên sheet đầu tiên
	sheetName := f.GetSheetName(0)

	// Duyệt qua nhiều ô (VD: từ A1 đến E10)
	var finalContent strings.Builder
	for rowIndex := 1; rowIndex <= 10; rowIndex++ {
		for colIndex := 1; colIndex <= 5; colIndex++ {
			cellName, _ := excelize.CoordinatesToCellName(colIndex, rowIndex)

			// Lấy Rich Text từ ô hiện tại
			richText, err := f.GetCellRichText(sheetName, cellName)
			if err != nil {
				log.Printf("Lỗi khi đọc %s: %s\n", cellName, err)
				continue
			}

			// Chuyển sang HTML
			htmlContent := richTextToHTML(richText)

			// Ghi dữ liệu vào file (kèm tên ô)
			finalContent.WriteString(fmt.Sprintf("%s: %s\n", cellName, htmlContent))
		}
	}

	// Lưu nội dung vào file .txt
	err = saveToTextFile(outputFile, finalContent.String())
	if err != nil {
		log.Fatalf("Lỗi khi lưu file .txt: %s", err)
	}

	fmt.Printf("File '%s' đã được lưu thành công!\n", outputFile)
}
