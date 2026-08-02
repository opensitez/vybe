// vybe-test: go/errors_join_unwrap/errorf_wrap_preserves_chain
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
import "errors"
var ErrIO = errors.New("io")
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { err := fmt.Errorf("read: %w", ErrIO)
__check(fmt.Sprint(err.Error()), "read: io")
__check(fmt.Sprint(errors.Is(err, ErrIO)), "true") }
