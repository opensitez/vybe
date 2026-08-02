// vybe-test: go/errors_join_unwrap/errors_unwrap_chain_three_levels
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
import "errors"
var ErrLeaf = errors.New("leaf")
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { e1 := fmt.Errorf("l1: %w", ErrLeaf)
e2 := fmt.Errorf("l2: %w", e1)
__check(fmt.Sprint(errors.Is(errors.Unwrap(e2), e1)), "true") }
