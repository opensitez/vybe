// vybe-test: go/errors_join_unwrap/errors_as_on_wrapped_in_join
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
import "errors"
type myErr struct { msg string }
func (e myErr) Error() string { return e.msg }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { inner := myErr{msg: "inner"}
joined := errors.Join(fmt.Errorf("wrap: %w", inner))
var target myErr
__check(fmt.Sprint(errors.As(joined, &target)), "true")
__check(fmt.Sprint(target.msg), "inner") }
