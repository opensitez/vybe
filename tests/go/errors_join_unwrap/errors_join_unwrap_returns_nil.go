// vybe-test: go/errors_join_unwrap/errors_join_unwrap_returns_nil
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

func main() { joined := errors.Join(errors.New("a"), errors.New("b"))
__check(fmt.Sprint(errors.Unwrap(joined) == nil), "true") }
