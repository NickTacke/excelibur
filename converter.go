package main

import (
	"bytes"
	"encoding/binary"
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

	ole, err := parseOLE2(data)
	if err != nil {
		return fmt.Errorf("error parsing OLE2 file: %v", err)
	}

	fmt.Println(ole)

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

// BIFF Workbook structure
type biffWorkbook struct {
	sheets  []biffSheet
	sst     []string       // Shared String Table
	formats map[int]string // Number formats
	xfs     []biffXF       // Cell formats
}

// BIFF Worksheet structure
type biffSheet struct {
	name       string
	rows       map[uint16]biffRow
	dimensions biffDimensions
}

// BIFF Row structure
type biffRow struct {
	rowIndex uint16
	cells    []biffCell
}

// BIFF Cell structure
type biffCell struct {
	row      uint16
	col      uint16
	cellType string // "s"=string, "n"=number, "b"=boolean, "f"=formula, "e"=error
	strVal   string
	numVal   float64
	boolVal  bool
	formula  string
	xfIndex  uint16
}

// BIFF Dimensions structure
type biffDimensions struct {
	firstRow uint16
	lastRow  uint16
	firstCol uint16
	lastCol  uint16
}

// BIFF XF (Cell format) structure
type biffXF struct {
	fontIndex      uint16
	formatIndex    uint16
	cellProtection uint16
	alignment      byte
	rotation       byte
	borders        [4]byte
	colors         [4]byte
	backgroundFill byte
}

func parseOLE2(data []byte) (*ole2, error) {
	// Check if the file is big enough to contain an OLE2 header
	if len(data) < int(SECTOR_SIZE) {
		return nil, errors.New("not a valid XLS file (file too small)")
	}

	// Create a new OLE2 structure
	ole := &ole2{}

	// Parse the OLE2 header
	header := ole2Header{}
	reader := bytes.NewReader(data)
	if err := binary.Read(reader, binary.LittleEndian, &header); err != nil {
		return nil, fmt.Errorf("error reading OLE2 header: %v", err)
	}
	ole.header = header

	// Validate the OLE2 header
	if !bytes.Equal(header.signature[:], []byte(OLE2_SIGNATURE)) {
		return nil, errors.New("invalid OLE2 signature")
	}

	// Determine the sector size and mini sector size
	ole.sectorSize = int(1 << header.sectorShift)
	ole.miniSectorSize = int(1 << header.miniSectorShift)

	// Read FAT sectors
	ole.fat = make([]uint32, 0, header.numFatSectors*uint32(ole.sectorSize/4))

	for i := 0; i < 109; i++ {
		// If the FAT entry is whitespace, skip it
		if header.difat[i] == 0xFFFFFFFF {
			continue
		}

		// Read the FAT entry
		sectorData := getSector(data, int(header.difat[i]), ole.sectorSize)
		fatEntries := make([]uint32, ole.sectorSize/4)

		if err := binary.Read(bytes.NewReader(sectorData), binary.LittleEndian, &fatEntries); err != nil {
			return nil, fmt.Errorf("error reading FAT sector: %v", err)
		}
		ole.fat = append(ole.fat, fatEntries...)
	}

	return ole, nil
}

// Get a sector from the OLE2 file
func getSector(data []byte, sectorId int, sectorSize int) []byte {
	offset := 512 + sectorId*sectorSize
	end := offset + sectorSize
	if end > len(data) {
		end = len(data)
	}
	return data[offset:end]
}

func main() {
	xls := XLSReader{}

	err := xls.ConvertFile("./sample.xls", "./sample.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
