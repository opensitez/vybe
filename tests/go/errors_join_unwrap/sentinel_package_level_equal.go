// vybe-test: go/errors_join_unwrap/sentinel_package_level_equal
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
import "errors"
var ErrA = errors.New("fail")
var ErrB = ErrA
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(ErrA == ErrB), "true") }
