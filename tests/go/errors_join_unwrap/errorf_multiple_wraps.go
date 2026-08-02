// vybe-test: go/errors_join_unwrap/errorf_multiple_wraps
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

func main() { e1 := fmt.Errorf("a: %w", errors.New("b"))
e2 := fmt.Errorf("c: %w", e1)
__check(fmt.Sprint(errors.Is(e2, errors.New("b"))), "false") }
