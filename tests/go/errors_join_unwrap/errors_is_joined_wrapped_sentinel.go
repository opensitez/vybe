// vybe-test: go/errors_join_unwrap/errors_is_joined_wrapped_sentinel
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
import "errors"
var ErrBase = errors.New("base")
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { wrapped := fmt.Errorf("w: %w", ErrBase)
joined := errors.Join(wrapped, errors.New("other"))
__check(fmt.Sprint(errors.Is(joined, ErrBase)), "true") }
