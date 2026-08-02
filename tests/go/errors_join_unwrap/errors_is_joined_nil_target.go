// vybe-test: go/errors_join_unwrap/errors_is_joined_nil_target
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
import "errors"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { joined := errors.Join(errors.New("x"))
__check(fmt.Sprint(errors.Is(joined, nil)), "false") }
