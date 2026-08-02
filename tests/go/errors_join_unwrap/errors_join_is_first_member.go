// vybe-test: go/errors_join_unwrap/errors_join_is_first_member
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

func main() { a := errors.New("a")
b := errors.New("b")
joined := errors.Join(a, b)
__check(fmt.Sprint(errors.Is(joined, a)), "true") }
