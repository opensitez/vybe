// vybe-test: go/errors_package/errors_new_error_string
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

func main() { err := errors.New("file not found")
__check(fmt.Sprint(err.Error()), "file not found") }
