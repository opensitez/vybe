// vybe-test: go/errors_join_unwrap/sentinel_is_not_equality
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
import "errors"
var ErrSent = errors.New("sent")
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { other := fmt.Errorf("wrap: %w", ErrSent)
__check(fmt.Sprint(other == ErrSent), "false")
__check(fmt.Sprint(errors.Is(other, ErrSent)), "true") }
