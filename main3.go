package main

import (
	"fmt"
	"log"
	"os"
	"strconv"

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

func main() {
	// Tải biến môi trường từ file .env
	err := godotenv.Load()
	if err != nil {
		log.Fatal("Lỗi khi tải file .env")
	}

	// Lấy tên file từ ENV
	filePath := os.Getenv("EXCEL_FILE")
	if filePath == "" {
		log.Fatal("Vui lòng đặt biến môi trường 'EXCEL_FILE' trong file .env!")
	}

	// Lấy dòng và cột từ ENV (mặc định: đọc từ hàng 1-10, cột A-E)
	startRow := getEnvInt("START_ROW", 1)
	endRow := getEnvInt("END_ROW", 10)
	startCol := getEnvInt("START_COL", 1)
	endCol := getEnvInt("END_COL", 5)

	// Mở file Excel
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		log.Fatalf("Không thể mở file: %s", err)
	}
	defer f.Close()

	// Lấy tên sheet đầu tiên
	sheetName := f.GetSheetName(0)

	// Duyệt qua các dòng và cột
	for rowIndex := startRow; rowIndex <= endRow; rowIndex++ {
		for colIndex := startCol; colIndex <= endCol; colIndex++ {
			cellName, _ := excelize.CoordinatesToCellName(colIndex, rowIndex)

			// Lấy nội dung dưới dạng Rich Text
			richText, err := f.GetCellRichText(sheetName, cellName)
			if err != nil {
				log.Fatal(err)
			}

			// Kiểm tra nếu richText bị nil hoặc rỗng
			if richText == nil || len(richText) == 0 {
				// Nếu không có Rich Text, lấy giá trị bình thường
				value, err := f.GetCellValue(sheetName, cellName)
				if err != nil {
					log.Fatal(err)
				}
				fmt.Printf("Cell %s: '%s' (No Rich Text)\n", cellName, value)
				continue
			}

			// Nếu có Rich Text, in từng phần với định dạng
			fmt.Printf("Cell %s:\n", cellName)
			for _, rt := range richText {
				font := rt.Font
				if font == nil { // Kiểm tra nil trước khi truy cập
					fmt.Printf("  - Text: '%s' (No font info)\n", rt.Text)
					continue
				}
				fmt.Printf("  - Text: '%s', Font: %s, Size: %.1f, Bold: %v, Italic: %v, Underline: %v, Color: #%s\n",
					rt.Text, font.Family, font.Size, font.Bold, font.Italic, font.Underline, font.Color)
			}
		}
	}
}
