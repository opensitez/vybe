// vybe-test: go/errors_join_unwrap/errorf_wrap_with_formatting
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

func main() { err := fmt.Errorf("failed after %d retries: %w", 3, errors.New("timeout"))
__check(fmt.Sprint(err.Error()), "failed after 3 retries: timeout") }
