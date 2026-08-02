// vybe-test: go/bufio_scanner_extended/scanner_lines_preserves_inner_spaces
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

func main() { sc := bufio.NewScanner(strings.NewReader("a b\nc d"))
sc.Split(bufio.ScanLines)
sc.Scan()
__check(fmt.Sprint(sc.Text()), "a b") }
