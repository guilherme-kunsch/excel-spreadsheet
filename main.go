package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// vai tentar abrir o arquivo
func openSheet(filename string) (*excelize.File, error) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, err
	}
	fmt.Println("Planilha aberta")
	return f, nil
}

func readSheet(f *excelize.File, sheet string) {
	rows, err := f.GetRows(sheet)
	if err != nil {
		fmt.Println(err)
		return
	}

	for _, row := range rows {
		for _, col := range row {
			fmt.Printf("%-15s", col) //separar as colunas com 15 caracteres
		}
		fmt.Println()
	}
}

func insertRow(f *excelize.File, sheet string, row int) error {
	if err := f.InsertRows(sheet, row, 1); err != nil {
		return err
	}
	fmt.Printf("Linha %d inserida na planilha %s\n", row, sheet)
	return nil
}

func writeCell(f *excelize.File, sheet string, cell string, value interface{}) error {
	if err := f.SetCellValue(sheet, cell, value); err != nil {
		return err
	}
	return nil
}

func saveFile(f *excelize.File, path string) error {
	if err := f.SaveAs(path); err != nil {
		return fmt.Errorf("failed to save file: %v", err)
	}
	return nil
}

func main() {
	f, err := openSheet("Pasta1.xlsx")
	if err != nil {
		fmt.Println("Erro ao abrir a planilha:", err)
		return
	}
	defer f.Close()

	if err := saveFile(f, "Pasta1.xlsx"); err != nil {
		fmt.Println(err)
	} else {
		fmt.Println("File saved successfully.")
	}

	readSheet(f, "Planilha1")

	writeCell(f, "Planilha1", "A6", "Teste")
	writeCell(f, "Planilha1", "B6", "Teste")
	writeCell(f, "Planilha1", "C6", "21")
	writeCell(f, "Planilha1", "D6", "10000")

	if err := insertRow(f, "Planilha1", 6); err != nil {
		fmt.Println("Erro ao inserir linha:", err)
		return
	}
}
