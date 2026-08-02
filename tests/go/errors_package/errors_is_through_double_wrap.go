// vybe-test: go/errors_package/errors_is_through_double_wrap
// origin: languages/go/tests/go/test_errors_package.rs

package main
import "fmt"
import "errors"
var ErrBase = errors.New("base")
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { err := fmt.Errorf("layer1: %w", fmt.Errorf("layer2: %w", ErrBase))
__check(fmt.Sprint(errors.Is(err, ErrBase)), "true") }
