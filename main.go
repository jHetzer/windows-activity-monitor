package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"syscall"
	"time"
	"unsafe"

	"github.com/Microsoft/go-winio/pkg/process"
	"github.com/bi-zone/go-fileversion"
	"golang.org/x/sys/windows"
)

var (
	mod                          = windows.NewLazyDLL("user32.dll")
	procGetWindowText            = mod.NewProc("GetWindowTextW")
	procGetWindowTextLength      = mod.NewProc("GetWindowTextLengthW")
	procGetWindowThreadProcessId = mod.NewProc("GetWindowThreadProcessId")
)

type (
	HANDLE uintptr
	HWND   HANDLE
)

func GetWindowTextLength(hwnd HWND) int {
	ret, _, _ := procGetWindowTextLength.Call(
		uintptr(hwnd))

	return int(ret)
}

func GetProcessProductName(path string) string {
	f, err := fileversion.New(path)
	if err != nil {
		log.Fatal(err)
	}
	return f.ProductName()
}

func GetProcessPath(pid uint32) string {

	// https://github.com/microsoft/hcsshim/blob/main/internal/uvm/stats.go
	// https://github.com/microsoft/hcsshim/blob/main/LICENSE (MIT)

	p, err := windows.OpenProcess(windows.PROCESS_QUERY_LIMITED_INFORMATION|windows.PROCESS_VM_READ, false, pid)
	if err != nil {
		log.Fatalln(err)
	}
	defer func(openedProcess windows.Handle) {
		// If we don't return this process handle, close it so it doesn't leak.
		if p == 0 {
			windows.Close(openedProcess)
		}
	}(p)
	// Querying vmmem's image name as a win32 path returns ERROR_GEN_FAILURE
	// for some reason, so we query it as an NT path instead.
	name, err := process.QueryFullProcessImageName(p, process.ImageNameFormatWin32Path)
	if err != nil {
		log.Fatalln(err)
	}
	return name
}

func GetWindowThreadProcessId(hwnd HWND) uintptr {
	var prcsId uintptr = 0
	procGetWindowThreadProcessId.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&prcsId)))

	return prcsId
}

func GetWindowText(hwnd HWND) string {
	textLen := GetWindowTextLength(hwnd) + 1

	buf := make([]uint16, textLen)
	procGetWindowText.Call(
		uintptr(hwnd),
		uintptr(unsafe.Pointer(&buf[0])),
		uintptr(textLen))

	return syscall.UTF16ToString(buf)
}

func getWindow(funcName string) uintptr {
	proc := mod.NewProc(funcName)
	hwnd, _, _ := proc.Call()
	return hwnd
}

var (
	//intFlag int
	//boolFlag bool
	path string
)

func main() {
	//flag.IntVar(&intFlag, "int", 1234, "help message")
	//flag.BoolVar(&boolFlag, "bool", false, "help message")
	flag.StringVar(&path, "path", "./", "Basepath for CSV-Reports")

	flag.Parse()

	fmt.Println("Report will be saved at: ", path)

	var lastString string
	var t time.Time
	timeformat := "20060102"

	filename := "records_" + time.Now().Format(timeformat) + ".csv"
	file, err := os.Create(filepath.Join(path, filename))

	if err != nil {
		log.Fatalln("failed to open file", err)
	}
	defer file.Close()

	w := csv.NewWriter(file)
	header := []string{
		"DateTime",
		"PID",
		"ProductName",
		"Info",
	}
	if err := w.Write(header); err != nil {
		log.Fatalln("error writing record to file", err)
	}

	for {
		if hwnd := getWindow("GetForegroundWindow"); hwnd != 0 {
			text := GetWindowText(HWND(hwnd))
			currentString := "window :" + text + "# hwnd:" + fmt.Sprint(hwnd)
			currentFileName := "records_" + time.Now().Format(timeformat) + ".csv"
			if filename != currentFileName {
				w.Flush()
				file.Close()
				filename = currentFileName
				file, err = os.Create(filepath.Join(path, filename))
				if err != nil {
					log.Fatalln("failed to open file", err)
				}
				w = csv.NewWriter(file)
				if err := w.Write(header); err != nil {
					log.Fatalln("error writing record to file", err)
				}
			}
			if currentString != lastString {
				t = time.Now()

				prPid := GetWindowThreadProcessId(HWND(hwnd))
				pid := fmt.Sprint(prPid)
				processPath := GetProcessPath(uint32(prPid))
				processProduct := GetProcessProductName(processPath)

				lastString = currentString

				row := []string{
					t.Format(time.RFC3339),
					pid,
					processProduct,
					text,
				}
				fmt.Println(row)
				if err := w.Write(row); err != nil {
					log.Fatalln("error writing record to file", err)
				}
				w.Flush()

			}
		}
	}
}
