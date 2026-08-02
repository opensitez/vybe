// vybe-test: go/errors_join_unwrap/errors_is_on_joined_wrapped_chain
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
import "errors"
var ErrDeep = errors.New("deep")
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { wrapped := fmt.Errorf("layer: %w", ErrDeep)
joined := errors.Join(errors.New("shallow"), wrapped)
__check(fmt.Sprint(errors.Is(joined, ErrDeep)), "true") }
