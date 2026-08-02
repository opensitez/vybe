// vybe-test: go/fmt_errors_print/fprintln_adds_newline
// origin: languages/go/tests/go/test_fmt_errors_print.rs

package main
import "fmt"
import "bytes"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var buf bytes.Buffer
fmt.Fprintln(&buf, "go")
__check(fmt.Sprint(buf.String()), "go\n") }
