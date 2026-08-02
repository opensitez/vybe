// vybe-test: go/errors_join_unwrap/errors_is_joined_sentinel_first
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
import "errors"
var ErrOne = errors.New("one")
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { joined := errors.Join(ErrOne, errors.New("two"))
__check(fmt.Sprint(errors.Is(joined, ErrOne)), "true") }
