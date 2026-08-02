// vybe-test: go/errors_package/errors_new_value_not_nil
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

func main() { err := errors.New("fail")
__check(fmt.Sprint(err != nil), "true") }
