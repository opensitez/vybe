// vybe-test: go/bufio_io/scanner_default_line_first_token
// origin: languages/go/tests/go/test_bufio_io.rs

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

func main() { sc := bufio.NewScanner(strings.NewReader("alpha\nbeta"))
sc.Scan()
__check(fmt.Sprint(sc.Text()), "alpha") }
