// vybe-test: go/errors_package/fmt_errorf_wrap_includes_cause_text
// origin: languages/go/tests/go/test_errors_package.rs

package main
import "fmt"
import "errors"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { err := fmt.Errorf("read failed: %w", errors.New("EOF"))
__check(fmt.Sprint(err.Error()), "read failed: EOF") }
