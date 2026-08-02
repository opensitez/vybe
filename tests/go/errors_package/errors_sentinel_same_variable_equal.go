// vybe-test: go/errors_package/errors_sentinel_same_variable_equal
// origin: languages/go/tests/go/test_errors_package.rs

package main
import "fmt"
import "errors"
var ErrDone = errors.New("done")
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(ErrDone == ErrDone), "true") }
