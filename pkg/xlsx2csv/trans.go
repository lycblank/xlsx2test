package xlsx2csv

import "github.com/lycblank/xlsx2text/pkg/xlsx2text"

func TransDir(dirName string, dstDirName string) error {
	return xlsx2text.TransDir(dirName, dstDirName, ",", ".csv")
}