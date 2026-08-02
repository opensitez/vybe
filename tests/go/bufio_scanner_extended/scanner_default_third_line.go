// vybe-test: go/bufio_scanner_extended/scanner_default_third_line
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

func main() { sc := bufio.NewScanner(strings.NewReader("1\n2\n3"))
sc.Scan()
sc.Scan()
sc.Scan()
__check(fmt.Sprint(sc.Text()), "3") }
