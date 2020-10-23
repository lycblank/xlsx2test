package main

import (
	"fmt"
	"github.com/lycblank/xlsx2text/pkg/xlsx2csv"
	"gopkg.in/alecthomas/kingpin.v2"
	"os"
)

var (
	srcDirName = kingpin.Flag("src", "src dir name").Short('s').String()
	dstDirName = kingpin.Flag("dst", "dst dir name").Short('d').String()
)

func main() {
	kingpin.Parse()
	err := xlsx2csv.TransDir(*srcDirName, *dstDirName)
	if err != nil {
		fmt.Printf("xlsx2csv failed. src:%s dst:%s err:%s\n", *srcDirName, *dstDirName, err)
		os.Exit(1)
	}
}



