// vybe-test: go/errors_join_unwrap/sentinel_wrapped_not_equal_unwrapped
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

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

func main() { w := fmt.Errorf("w: %w", ErrRoot)
__check(fmt.Sprint(w == ErrRoot), "false") }
