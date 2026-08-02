// vybe-test: go/errors_join_unwrap/errors_unwrap_joined_not_chain
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

func main() { inner := errors.New("inner")
joined := errors.Join(fmt.Errorf("wrap: %w", inner))
__check(fmt.Sprint(errors.Is(joined, inner)), "true") }
