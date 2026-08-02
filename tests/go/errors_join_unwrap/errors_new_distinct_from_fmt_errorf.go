// vybe-test: go/errors_join_unwrap/errors_new_distinct_from_fmt_errorf
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

func main() { a := errors.New("msg")
b := fmt.Errorf("msg")
__check(fmt.Sprint(a == b), "false")
__check(fmt.Sprint(a.Error() == b.Error()), "true") }
