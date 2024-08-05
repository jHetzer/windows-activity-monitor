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

	modOleacc                    = windows.NewLazyDLL("Oleacc.dll")
	procGetProcessHandleFromHwnd = modOleacc.NewProc("GetProcessHandleFromHwnd")
)

/*
HANDLE WINAPI GetProcessHandleFromHwnd(
  _In_ HWND hwnd
);
*/

type (
	HANDLE uintptr
	HWND   HANDLE
)

func GetWindowTextLength(hwnd HWND) int {
	ret, _, _ := procGetWindowTextLength.Call(
		uintptr(hwnd),
	)

	return int(ret)
}

func GetProcessProductName(path string) (string, error) {
	f, err := fileversion.New(path)
	if err != nil {
		panic(err)
	}
	return f.ProductName(), nil
}

func GetProcessPath(hwnd HWND) (string, error) {
	//func GetProcessPath(pid uint32) string {

	ret, _, _ := procGetProcessHandleFromHwnd.Call(
		uintptr(hwnd),
	)

	procHandle := windows.Handle(ret)
	name, err := process.QueryFullProcessImageName(procHandle, process.ImageNameFormatWin32Path)

	return name, err
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
					panic(err)
				}
				w = csv.NewWriter(file)
				if err := w.Write(header); err != nil {
					panic(err)
				}
			}
			if currentString != lastString {
				lastString = currentString
				t = time.Now()

				prPid := GetWindowThreadProcessId(HWND(hwnd))
				pid := fmt.Sprint(prPid)
				processPath, pathErr := GetProcessPath(HWND(hwnd))
				var processProduct string
				if pathErr != nil {
					processProduct = pathErr.Error()
					//processPath = "File path not accessable:" + err.Error()
				} else {
					processProduct, _ = GetProcessProductName(processPath)
				}

				row := []string{
					t.Format(time.RFC3339),
					pid,
					processProduct,
					text,
				}
				fmt.Println(row)
				if err := w.Write(row); err != nil {

					panic(err)
				}
				w.Flush()
			}
		}
	}
}
