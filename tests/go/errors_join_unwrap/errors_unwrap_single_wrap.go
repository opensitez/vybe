// vybe-test: go/errors_join_unwrap/errors_unwrap_single_wrap
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
import "errors"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { inner := errors.New("core")
outer := fmt.Errorf("wrap: %w", inner)
__check(fmt.Sprint(errors.Unwrap(outer).Error()), "core") }
