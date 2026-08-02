// vybe-test: go/errors_package/errors_is_nil_target_always_false
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

func main() { __check(fmt.Sprint(errors.Is(errors.New("x"), nil)), "false") }
