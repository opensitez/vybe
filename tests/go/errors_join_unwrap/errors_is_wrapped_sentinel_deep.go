// vybe-test: go/errors_join_unwrap/errors_is_wrapped_sentinel_deep
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
import "errors"
var ErrBottom = errors.New("bottom")
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { e := fmt.Errorf("a: %w", fmt.Errorf("b: %w", fmt.Errorf("c: %w", ErrBottom)))
__check(fmt.Sprint(errors.Is(e, ErrBottom)), "true") }
