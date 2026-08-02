// vybe-test: go/errors_package/errors_unwrap_join_returns_nil
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

func main() { joined := errors.Join(errors.New("a"), errors.New("b"))
__check(fmt.Sprint(errors.Unwrap(joined) == nil), "true") }
