package main

import (
	"bytes"
	"errors"
	"fmt"
	"os"
)

const (
	// Signatures
	OLE2_SIGNATURE = "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"

	// OLE2 Sector Sizes
	SECTOR_SIZE      = 512
	MINI_SECTOR_SIZE = 64

	// BIFF Record Types
	RECORD_BOF        = 0x0809 // Beginning of File
	RECORD_EOF        = 0x000A // End of File
	RECORD_BOUNDSHEET = 0x0085 // Worksheet information
	RECORD_SST        = 0x00FC // Shared String Table
	RECORD_LABELSST   = 0x00FD // Cell with string from SST
	RECORD_NUMBER     = 0x0203 // Cell with number data
	RECORD_FORMULA    = 0x0006 // Cell with formula
	RECORD_STRING     = 0x0207 // Result of string formula
	RECORD_BOOLERR    = 0x0205 // Cell with boolean or error
	RECORD_FORMAT     = 0x041E // Number format definition
	RECORD_XF         = 0x00E0 // Cell formatting
	RECORD_MULRK      = 0x00BD // Multiple RK cells
	RECORD_RK         = 0x027E // Cell with RK value
	RECORD_BLANK      = 0x0201 // Empty cell
	RECORD_ROW        = 0x0208 // Row properties
)

type XLSReader struct {
	// Configuration parameters
	IgnoreErrors bool
	Debug        bool
}

func (xls *XLSReader) ConvertFile(xlsIn string, xlsxOut string) error {
	// Read the entire file into memory
	data, err := os.ReadFile(xlsIn)
	if err != nil {
		return err
	}

	// Check if it's an OLE2 file
	if !bytes.HasPrefix(data, []byte(OLE2_SIGNATURE)) {
		return errors.New("not a valid XLS file (OLE2 signature not found)")
	}

	fmt.Print(data)

	return nil
}

// OLE2 Header structure
type ole2Header struct {
	signature        [8]byte
	clsid            [16]byte
	minorVersion     uint16
	majorVersion     uint16
	byteOrder        uint16
	sectorShift      uint16
	miniSectorShift  uint16
	reserved         [6]byte
	numDirSectors    uint32
	numFatSectors    uint32
	firstDirSector   uint32
	transactionSig   uint32
	miniStreamCutoff uint32
	firstMiniFatSec  uint32
	numMiniFatSecs   uint32
	firstDifatSec    uint32
	numDifatSecs     uint32
	difat            [109]uint32
}

// Directory entry structure
type dirEntry struct {
	name          string
	nameRaw       [64]byte
	entryType     byte
	colorFlag     byte
	leftSibID     uint32
	rightSibID    uint32
	childID       uint32
	clsid         [16]byte
	stateBits     uint32
	createTime    uint64
	modifyTime    uint64
	startSector   uint32
	streamSize    uint64
	isDirectory   bool
	isRootStorage bool
}

// OLE2 structure
type ole2 struct {
	header         ole2Header
	sectorSize     int
	miniSectorSize int
	dirEntries     []dirEntry
	fat            []uint32
	miniFat        []uint32
	sectors        [][]byte
	miniSectors    [][]byte
}

func main() {
	xls := XLSReader{}

	err := xls.ConvertFile("./sample.xls", "./sample.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
