// vybe-test: go/errors_join_unwrap/sentinel_joined_same_var
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
import "errors"
var ErrX = errors.New("x")
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { joined := errors.Join(ErrX, errors.New("y"))
__check(fmt.Sprint(errors.Is(joined, ErrX)), "true") }
