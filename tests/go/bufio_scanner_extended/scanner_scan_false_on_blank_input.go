// vybe-test: go/bufio_scanner_extended/scanner_scan_false_on_blank_input
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs

package main
import "fmt"
import "bufio"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { sc := bufio.NewScanner(strings.NewReader(""))
sc.Split(bufio.ScanWords)
__check(fmt.Sprint(sc.Scan()), "false") }
