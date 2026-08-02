// vybe-test: go/errors_package/errors_wrapped_not_equal_but_is_true
// origin: languages/go/tests/go/test_errors_package.rs

package main
import "fmt"
import "errors"
var ErrSentinel = errors.New("sentinel")
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { wrapped := fmt.Errorf("wrap: %w", ErrSentinel)
__check(fmt.Sprint(wrapped == ErrSentinel), "false")
__check(fmt.Sprint(errors.Is(wrapped, ErrSentinel)), "true") }
