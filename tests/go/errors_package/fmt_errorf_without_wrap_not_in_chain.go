// vybe-test: go/errors_package/fmt_errorf_without_wrap_not_in_chain
// origin: languages/go/tests/go/test_errors_package.rs

package main
import "fmt"
import "errors"
var ErrRoot = errors.New("root")
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { err := fmt.Errorf("outer: %v", ErrRoot)
__check(fmt.Sprint(errors.Is(err, ErrRoot)), "false") }
