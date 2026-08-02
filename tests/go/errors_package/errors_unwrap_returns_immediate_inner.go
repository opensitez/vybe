// vybe-test: go/errors_package/errors_unwrap_returns_immediate_inner
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

func main() { inner := errors.New("inner")
outer := fmt.Errorf("outer: %w", inner)
__check(fmt.Sprint(errors.Unwrap(outer) == inner), "true") }
