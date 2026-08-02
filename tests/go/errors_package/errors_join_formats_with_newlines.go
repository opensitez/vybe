// vybe-test: go/errors_package/errors_join_formats_with_newlines
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

func main() { err := errors.Join(errors.New("first"), errors.New("second"))
__check(fmt.Sprint(err.Error()), "first\nsecond") }
