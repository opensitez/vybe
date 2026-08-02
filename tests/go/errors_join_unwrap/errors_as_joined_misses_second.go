// vybe-test: go/errors_join_unwrap/errors_as_joined_misses_second
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
import "errors"
type coded struct { code int }
func (c coded) Error() string { return "coded" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { joined := errors.Join(errors.New("plain"), coded{code: 1})
var target coded
__check(fmt.Sprint(errors.As(joined, &target)), "true") }
