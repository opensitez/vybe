// vybe-test: go/errors_package/errors_as_misses_unrelated_type
// origin: languages/go/tests/go/test_errors_package.rs

package main
import "fmt"
import "errors"
type coded struct { n int }
func (c coded) Error() string { return "coded" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { err := errors.New("plain")
var target coded
__check(fmt.Sprint(errors.As(err, &target)), "false") }
