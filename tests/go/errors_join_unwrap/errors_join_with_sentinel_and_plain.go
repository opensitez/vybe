// vybe-test: go/errors_join_unwrap/errors_join_with_sentinel_and_plain
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
import "errors"
var ErrFatal = errors.New("fatal")
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { err := errors.Join(ErrFatal, errors.New("warning"))
__check(fmt.Sprint(errors.Is(err, ErrFatal)), "true")
__check(fmt.Sprint(errors.Is(err, errors.New("warning"))), "false") }
