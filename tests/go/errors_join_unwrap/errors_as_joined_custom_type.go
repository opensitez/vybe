// vybe-test: go/errors_join_unwrap/errors_as_joined_custom_type
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
import "errors"
type coded struct { code int }
func (c coded) Error() string { return fmt.Sprintf("code %d", c.code) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { inner := coded{code: 42}
joined := errors.Join(inner, errors.New("plain"))
var target coded
__check(fmt.Sprint(errors.As(joined, &target)), "true")
__check(fmt.Sprint(target.code), "42") }
