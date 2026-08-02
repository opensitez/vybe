// vybe-test: go/errors_package/errors_is_matches_sentinel
// origin: languages/go/tests/go/test_errors_package.rs

package main
import "fmt"
import "errors"
var ErrNotFound = errors.New("not found")
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { err := fmt.Errorf("open: %w", ErrNotFound)
__check(fmt.Sprint(errors.Is(err, ErrNotFound)), "true") }
