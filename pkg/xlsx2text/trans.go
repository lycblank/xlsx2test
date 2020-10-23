package xlsx2text

import (
	"bufio"
	"fmt"
	"os"
	"path"
	"path/filepath"
	"strings"

	"github.com/tealeg/xlsx/v3"
)

func TransDir(dirName string, dstDirName string, delimiter string, ext string) error {
	absDirName, err := filepath.Abs(dirName)
	if err != nil {
		return err
	}
	dstPrefix := filepath.Base(absDirName)
	err = filepath.Walk(dirName, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if info == nil || info.IsDir() {
			return nil
		}
		fname := info.Name()
		if fname == "" || fname[0] == '~' { // 临时文件直接忽略掉
			return nil
		}
		exttmp := filepath.Ext(fname)
		if exttmp != ".xlsx" && exttmp != ".xls" { // 不是excel文件
			return nil
		}
		err = TransXlsx(path, dstDirName, dstPrefix, delimiter, ext)
		return err
	})
	return err
}

func TransXlsx(xlsxFileName string, dstDirName string, dstPrefix string, delimiter string, ext string) error {
	xlFile, err := xlsx.OpenFile(xlsxFileName)
	if err != nil {
		return err
	}
	for _, sheet := range xlFile.Sheets {
		if err := transXlsxSheet(sheet, xlsxFileName, dstDirName, dstPrefix, delimiter, ext); err != nil {
			return err
		}
	}
	return nil
}

func transXlsxSheet(sheet *xlsx.Sheet, xlsxFileName string, dstDirName string, dstPrefix string, delimiter string, ext string) error {
	sfname := path.Base(xlsxFileName)
	sfname = strings.Split(sfname, ".")[0]
	dstFileName := fmt.Sprintf("%s_%s_%s%s", dstPrefix, sfname, sheet.Name,ext)
	dstFileName = path.Join(dstDirName, dstFileName)
	dstFile, err := os.OpenFile(dstFileName, os.O_CREATE|os.O_TRUNC|os.O_WRONLY, 0666)
	if err != nil {
		return err
	}
	defer dstFile.Close()
	bufferWritten := bufio.NewWriter(dstFile)
	defer bufferWritten.Flush()
	err = sheet.ForEachRow(func(row *xlsx.Row) error {
		if row == nil {
			return nil
		}
		vals := make([]string, 0, 32)
		err := row.ForEachCell(func(cell *xlsx.Cell) error {
			str, err := cell.FormattedValue()
			if err != nil {
				return err
			}
			vals = append(vals, str)
			return nil
		})
		if err != nil {
			return err
		}
		if len(vals) > 0 {
			bufferWritten.WriteString(strings.Join(vals, delimiter))
			bufferWritten.WriteByte('\n')
		}
		return nil
	})
	if err != nil {
		return err
	}
	return nil
}


