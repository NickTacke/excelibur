package main

import (
	"bytes"
	"encoding/binary"
	"errors"
	"fmt"
	"os"
	"strings"
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

	// Print the ole2 headers
	fmt.Println("XLS File Information:")
	fmt.Println("---------------------")
	fmt.Printf("File Signature: %X\n", ole.header.Signature)
	fmt.Printf("Minor Version: %d\n", ole.header.MinorVersion)
	fmt.Printf("Major Version: %d\n", ole.header.DllVersion)
	fmt.Printf("Byte Order: 0x%04X\n", ole.header.ByteOrder)
	fmt.Printf("Sector Size: %d bytes (2^%d)\n", 1<<ole.header.SectorShift, ole.header.SectorShift)
	fmt.Printf("Mini Sector Size: %d bytes (2^%d)\n", 1<<ole.header.MiniSectorShift, ole.header.MiniSectorShift)
	fmt.Printf("Number of FAT Sectors: %d\n", ole.header.NumFATSectors)
	fmt.Printf("First Directory Sector: %d\n", ole.header.FirstDirSector)
	fmt.Printf("Mini Stream Cutoff: %d bytes\n", ole.header.MiniStreamCutoff)

	// Get the workbook stream
	workbookStream, err := ole.getStream("Workbook")
	if err != nil {
		workbookStream, err = ole.getStream("Book")
		if err != nil {
			return fmt.Errorf("error getting workbook stream: %v", err)
		}
	}

	// Print the raw workbook stream
	fmt.Printf("Workbook Stream:\n")
	for _, b := range workbookStream {
		fmt.Printf("%02X ", b)
	}

	return nil
}

// OLE2 Header structure
type Ole2Header struct {
	Signature        [8]byte     // Should be D0 CF 11 E0 A1 B1 1A E1
	CLSID            [16]byte    // Class ID (usually all zeros)
	MinorVersion     uint16      // Minor version of the format
	DllVersion       uint16      // Major version of the format
	ByteOrder        uint16      // Byte order (0xFFFE for little-endian)
	SectorShift      uint16      // Power of 2, sector size is 2^SectorShift (usually 9, for 512 bytes)
	MiniSectorShift  uint16      // Power of 2, mini-sector size is 2^MiniSectorShift (usually 6, for 64 bytes)
	Reserved1        [6]byte     // Reserved, must be zero
	NumDirSectors    uint32      // Number of directory sectors (usually 0 for < 4 MB files)
	NumFATSectors    uint32      // Number of FAT sectors
	FirstDirSector   uint32      // First directory sector location
	TransactionSig   uint32      // Transaction signature number
	MiniStreamCutoff uint32      // Maximum size for mini-stream (usually 4096 bytes)
	FirstMiniFATSec  uint32      // First mini-FAT sector location
	NumMiniFATSecs   uint32      // Number of mini-FAT sectors
	FirstDIFATSec    uint32      // First DIFAT sector location
	NumDIFATSecs     uint32      // Number of DIFAT sectors
	DIFAT            [109]uint32 // First 109 DIFAT entries
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
	header         Ole2Header
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
	header := Ole2Header{}
	reader := bytes.NewReader(data)
	if err := binary.Read(reader, binary.LittleEndian, &header); err != nil {
		return nil, fmt.Errorf("error reading OLE2 header: %v", err)
	}
	ole.header = header

	// Validate the OLE2 header
	if !bytes.Equal(header.Signature[:], []byte(OLE2_SIGNATURE)) {
		return nil, errors.New("invalid OLE2 signature")
	}

	// Determine the sector size and mini sector size
	ole.sectorSize = int(1 << header.SectorShift)
	ole.miniSectorSize = int(1 << header.MiniSectorShift)

	// Read FAT sectors
	ole.fat = make([]uint32, 0, header.NumFATSectors*uint32(ole.sectorSize/4))
	for i := 0; i < 109; i++ {
		// If the FAT entry is whitespace, skip it
		if header.DIFAT[i] == 0xFFFFFFFF {
			continue
		}

		// Read the FAT entry
		sectorData := getSector(data, int(header.DIFAT[i]), ole.sectorSize)
		fatEntries := make([]uint32, ole.sectorSize/4)

		if err := binary.Read(bytes.NewReader(sectorData), binary.LittleEndian, &fatEntries); err != nil {
			return nil, fmt.Errorf("error reading FAT sector: %v", err)
		}
		ole.fat = append(ole.fat, fatEntries...)
	}

	// Parse sectors
	numSectors := (len(data) - 512) / ole.sectorSize
	ole.sectors = make([][]byte, numSectors)
	for i := 0; i < numSectors; i++ {
		ole.sectors[i] = getSector(data, i, ole.sectorSize)
	}

	// Read directory entries
	dirSectors := readChain(ole.fat, int(header.FirstDirSector))
	dirData := make([]byte, 0, len(dirSectors)*ole.sectorSize)
	for _, sectorId := range dirSectors {
		if int(sectorId) >= len(ole.sectors) {
			return nil, fmt.Errorf("invalid directory sector ID in directory chain: %d", sectorId)
		}
		dirData = append(dirData, ole.sectors[sectorId]...)
	}

	// Parse the directory entries
	for i := 0; i < len(dirData); i += 128 {
		// Check if the directory is valid
		if i+128 > len(dirData) {
			break
		}

		// Read the directory entry
		entry := dirEntry{}
		entryReader := bytes.NewReader(dirData[i : i+128])

		// Read the entry name
		if err := binary.Read(entryReader, binary.LittleEndian, &entry.nameRaw); err != nil {
			return nil, fmt.Errorf("error reading directory entry name: %v", err)
		}

		// Convert the name to a string
		nameBuffer := bytes.NewBuffer(nil)
		for j := 0; j < 32; j++ {
			wchar := binary.LittleEndian.Uint16(entry.nameRaw[j*2 : j*2+2])
			if wchar == 0 {
				break
			}
			nameBuffer.WriteRune(rune(wchar))
		}
		entry.name = nameBuffer.String()

		// Read the entry type
		if err := binary.Read(entryReader, binary.LittleEndian, &entry.entryType); err != nil {
			return nil, fmt.Errorf("error reading directory entry type: %v", err)
		}

		// Read the entry color flag
		if err := binary.Read(entryReader, binary.LittleEndian, &entry.colorFlag); err != nil {
			return nil, fmt.Errorf("error reading directory entry color flag: %v", err)
		}

		// Read the entry left sibling ID
		if err := binary.Read(entryReader, binary.LittleEndian, &entry.leftSibID); err != nil {
			return nil, fmt.Errorf("error reading directory entry left sibling ID: %v", err)
		}

		// Read the entry right sibling ID
		if err := binary.Read(entryReader, binary.LittleEndian, &entry.rightSibID); err != nil {
			return nil, fmt.Errorf("error reading directory entry right sibling ID: %v", err)
		}

		// Read the entry child ID
		if err := binary.Read(entryReader, binary.LittleEndian, &entry.childID); err != nil {
			return nil, fmt.Errorf("error reading directory entry child ID: %v", err)
		}

		// Read the entry class ID
		if err := binary.Read(entryReader, binary.LittleEndian, &entry.clsid); err != nil {
			return nil, fmt.Errorf("error reading directory entry class ID: %v", err)
		}

		// Read the entry state bits
		if err := binary.Read(entryReader, binary.LittleEndian, &entry.stateBits); err != nil {
			return nil, fmt.Errorf("error reading directory entry state bits: %v", err)
		}

		// Read the entry create time
		if err := binary.Read(entryReader, binary.LittleEndian, &entry.createTime); err != nil {
			return nil, fmt.Errorf("error reading directory entry create time: %v", err)
		}

		// Read the entry modify time
		if err := binary.Read(entryReader, binary.LittleEndian, &entry.modifyTime); err != nil {
			return nil, fmt.Errorf("error reading directory entry modify time: %v", err)
		}

		// Read the entry start sector
		if err := binary.Read(entryReader, binary.LittleEndian, &entry.startSector); err != nil {
			return nil, fmt.Errorf("error reading directory entry start sector: %v", err)
		}

		// Read the entry stream size
		if err := binary.Read(entryReader, binary.LittleEndian, &entry.streamSize); err != nil {
			return nil, fmt.Errorf("error reading directory entry stream size: %v", err)
		}

		// TODO: Analyse the entry type
		// Print the entry type
		fmt.Printf("Type: %d\n", entry.entryType)

		// Set flags based on the entry type
		entry.isDirectory = entry.entryType == 1
		entry.isRootStorage = entry.name == "Root Entry"

		// Add the entry to the list
		ole.dirEntries = append(ole.dirEntries, entry)
	}

	// Find the root storage entry
	var rootStorage *dirEntry
	for i := range ole.dirEntries {
		if ole.dirEntries[i].isRootStorage {
			rootStorage = &ole.dirEntries[i]
			break
		}
	}

	// Check if the root storage entry exists
	if rootStorage == nil {
		return nil, errors.New("root storage not found")
	}

	// Read the Mini FAT if it exists
	if header.NumMiniFATSecs > 0 {
		miniFatChain := readChain(ole.fat, int(header.FirstMiniFATSec))
		miniFatData := make([]byte, 0, len(miniFatChain)*ole.sectorSize)
		for _, sectorId := range miniFatChain {
			miniFatData = append(miniFatData, ole.sectors[sectorId]...)
		}

		// Parse the Mini FAT
		ole.miniFat = make([]uint32, len(miniFatData)/4)
		if err := binary.Read(bytes.NewReader(miniFatData), binary.LittleEndian, &ole.miniFat); err != nil {
			return nil, fmt.Errorf("error reading Mini FAT: %v", err)
		}
	}

	// Read the mini stream
	miniStreamChain := readChain(ole.fat, int(rootStorage.startSector))
	miniStreamData := make([]byte, 0, len(miniStreamChain)*ole.sectorSize)
	for _, sectorId := range miniStreamChain {
		miniStreamData = append(miniStreamData, ole.sectors[sectorId]...)
	}

	// Truncate the mini stream to the mini stream cutoff
	if int(rootStorage.streamSize) < len(miniStreamData) {
		miniStreamData = miniStreamData[:int(rootStorage.streamSize)]
	}

	// Parse mini sectors
	numMiniSectors := len(miniStreamData) / ole.miniSectorSize
	ole.miniSectors = make([][]byte, numMiniSectors)
	for i := 0; i < numMiniSectors; i++ {
		start := i * ole.miniSectorSize
		end := start + ole.miniSectorSize
		if end > len(miniStreamData) {
			end = len(miniStreamData)
		}
		ole.miniSectors[i] = miniStreamData[start:end]
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

// Read a chain of sectors from the FAT
func readChain(fat []uint32, startSector int) []int {
	if startSector < 0 || startSector >= len(fat) || startSector == 0xFFFFFFFE {
		return []int{}
	}

	chain := []int{startSector}
	nextSector := int(fat[startSector])

	// Loop over the sectors
	for nextSector != 0xFFFFFFFE {
		if nextSector < 0 || nextSector >= len(fat) {
			break
		}

		chain = append(chain, nextSector)
		nextSector = int(fat[nextSector])
	}

	return chain
}

// Get the stream data from an OLE2 structure
func (ole *ole2) getStream(name string) ([]byte, error) {
	// Find the directory entry with the given name
	var streamEntry *dirEntry
	for i := range ole.dirEntries {
		if strings.EqualFold(ole.dirEntries[i].name, name) {
			streamEntry = &ole.dirEntries[i]
			break
		}
	}
	if streamEntry == nil {
		return nil, fmt.Errorf("stream not found: %s", name)
	}

	// Check if this is a mini stream or a regular stream
	if streamEntry.streamSize < uint64(ole.header.MiniStreamCutoff) {
		// Read from the mini stream
		miniChain := readChain(ole.miniFat, int(streamEntry.startSector))
		data := make([]byte, 0, streamEntry.streamSize)
		for _, sectorId := range miniChain {
			if sectorId >= len(ole.miniSectors) {
				return nil, fmt.Errorf("invalid mini sector ID: %d", sectorId)
			}
			data = append(data, ole.miniSectors[sectorId]...)
		}
		return data[:streamEntry.streamSize], nil
	} else {
		// Read from the regular stream
		chain := readChain(ole.fat, int(streamEntry.startSector))
		data := make([]byte, 0, streamEntry.streamSize)
		for _, sectorId := range chain {
			if sectorId >= len(ole.sectors) {
				return nil, fmt.Errorf("invalid sector ID: %d", sectorId)
			}
			data = append(data, ole.sectors[sectorId]...)
		}
		return data[:streamEntry.streamSize], nil
	}
}

func main() {
	// Convert the sample file
	xls := XLSReader{}
	err := xls.ConvertFile("./sample.xls", "./sample.xlsx")
	if err != nil {
		fmt.Println(err)
	}
}
