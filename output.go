package main

import (
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

// Chuyển Rich Text từ Excel thành HTML, giữ tất cả nội dung trong một <p> và thay \n bằng <br>
func richTextToHTML(richText []excelize.RichTextRun) string {
	var result strings.Builder
	result.WriteString("<p>")

	for i, rt := range richText {
		text := rt.Text
		if rt.Font != nil {
			if rt.Font.Bold {
				text = "<b>" + text + "</b>"
			}
			if rt.Font.Italic {
				text = "<i>" + text + "</i>"
			}
			if rt.Font.Underline == "single" {
				text = "<u>" + text + "</u>"
			}
		}
		if i > 0 {
			result.WriteString(" ")
		}
		result.WriteString(strings.TrimSpace(strings.ReplaceAll(text, "\n", "<br>")))
	}

	result.WriteString("</p>")
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

	// Lấy dòng và cột từ ENV
	startRow := getEnvInt("START_ROW", 1)
	endRow := getEnvInt("END_ROW", 10)
	startCol := getEnvInt("START_COL", 1)
	endCol := getEnvInt("END_COL", 5)

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

	// Duyệt qua các ô theo phạm vi đã chỉ định
	var finalContent strings.Builder
	for rowIndex := startRow; rowIndex <= endRow; rowIndex++ {
		for colIndex := startCol; colIndex <= endCol; colIndex++ {
			cellName, _ := excelize.CoordinatesToCellName(colIndex, rowIndex)

			// Lấy Rich Text từ ô hiện tại
			richText, err := f.GetCellRichText(sheetName, cellName)
			var content string

			if err == nil && richText != nil && len(richText) > 0 {
				content = richTextToHTML(richText)
			} else {
				value, err := f.GetCellValue(sheetName, cellName)
				if err != nil {
					log.Printf("Lỗi khi đọc %s: %s\n", cellName, err)
					continue
				}
				trimmedValue := strings.TrimSpace(value)
				if trimmedValue == "" {
					content = "<p>&nbsp;</p>"
				} else {
					content = "<p>" + strings.ReplaceAll(trimmedValue, "\n", "</p><p>") + "</p>"
				}
			}

			finalContent.WriteString(fmt.Sprintf("%s: %s\n", cellName, content))
		}
	}

	// Lưu nội dung vào file .txt
	err = saveToTextFile(outputFile, finalContent.String())
	if err != nil {
		log.Fatalf("Lỗi khi lưu file .txt: %s", err)
	}

	fmt.Printf("File '%s' đã được lưu thành công với phạm vi ô [%s] đến [%s]!\n", outputFile, 
		fmt.Sprintf("R%dC%d", startRow, startCol), fmt.Sprintf("R%dC%d", endRow, endCol))
}
