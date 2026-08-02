// vybe-test: go/errors_package/errors_is_rejects_unrelated_target
// origin: languages/go/tests/go/test_errors_package.rs

package main
import "fmt"
import "errors"
var ErrA = errors.New("a")
var ErrB = errors.New("b")
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(errors.Is(ErrA, ErrB)), "false") }
