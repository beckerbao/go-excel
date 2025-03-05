package main

import (
	"go-excel/common"
	"log"
	"os"
	"strconv"

	"github.com/joho/godotenv"
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

	// Lấy tên file Excel từ ENV
	sourceFile := os.Getenv("EXCEL_FILE")
	if sourceFile == "" {
		log.Fatal("Vui lòng đặt biến môi trường 'EXCEL_FILE' trong file .env!")
	}

	// Tên file Excel mới (tạo bản sao)
	// targetFile := "copy_" + sourceFile

	// Tên file TXT đầu ra
	outputTxtFile := "output.txt"

	// Tên file Excel sau khi import lại từ TXT
	importedExcelFile := "imported_" + sourceFile

	// // Lấy dòng và cột từ ENV (mặc định: đọc từ hàng 1-10, cột A-E)
	// startRow := getEnvInt("START_ROW", 1)
	// endRow := getEnvInt("END_ROW", 10)
	// startCol := getEnvInt("START_COL", 1)
	// endCol := getEnvInt("END_COL", 5)

	// // Mở file Excel gốc
	// srcExcel, err := excelize.OpenFile(sourceFile)
	// if err != nil {
	// 	log.Fatalf("Không thể mở file gốc: %s", err)
	// }
	// defer srcExcel.Close()

	// // Tạo file Excel mới
	// destExcel := excelize.NewFile()
	// sheetName := srcExcel.GetSheetName(0)
	// destExcel.NewSheet(sheetName)

	// // Chuỗi tổng hợp nội dung để lưu vào file .txt
	// var finalContent strings.Builder

	// // Duyệt từng ô để copy nội dung và định dạng
	// for rowIndex := startRow; rowIndex <= endRow; rowIndex++ {
	// 	for colIndex := startCol; colIndex <= endCol; colIndex++ {
	// 		cellName, _ := excelize.CoordinatesToCellName(colIndex, rowIndex)

	// 		richText, err := srcExcel.GetCellRichText(sheetName, cellName)
	// 		var cellContent string
	// 		if err == nil && richText != nil && len(richText) > 0 {
	// 			destExcel.SetCellRichText(sheetName, cellName, richText)
	// 			cellContent = common.RichTextToHTML(richText)
	// 		} else {
	// 			value, _ := srcExcel.GetCellValue(sheetName, cellName)
	// 			destExcel.SetCellValue(sheetName, cellName, value)
	// 			cellContent = value
	// 		}

	// 		// Lưu style
	// 		styleID, err := srcExcel.GetCellStyle(sheetName, cellName)
	// 		if err == nil {
	// 			destExcel.SetCellStyle(sheetName, cellName, cellName, styleID)
	// 		}

	// 		// Ghi dữ liệu vào file TXT
	// 		finalContent.WriteString(fmt.Sprintf("%s: %s\n", cellName, cellContent))
	// 	}
	// }

	// Lưu file Excel mới
	// _ = destExcel.SaveAs(targetFile)

	// Lưu nội dung vào file .txt
	// _ = common.SaveToTextFile(outputTxtFile, finalContent.String())

	// Chuyển file TXT thành Excel với Rich Text
	common.ImportFromTextToExcel(outputTxtFile, importedExcelFile)
}
