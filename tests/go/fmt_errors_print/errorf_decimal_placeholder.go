// vybe-test: go/fmt_errors_print/errorf_decimal_placeholder
// origin: languages/go/tests/go/test_fmt_errors_print.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { err := fmt.Errorf("status %d", 503)
__check(fmt.Sprint(err.Error()), "status 503") }
