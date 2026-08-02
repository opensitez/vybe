// vybe-test: go/fmt_errors_print/fprint_multiple_values_spaced
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
fmt.Fprint(&buf, "a", 1)
__check(fmt.Sprint(buf.String()), "a1") }
