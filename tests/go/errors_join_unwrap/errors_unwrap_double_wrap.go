// vybe-test: go/errors_join_unwrap/errors_unwrap_double_wrap
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

func main() { base := errors.New("base")
mid := fmt.Errorf("mid: %w", base)
outer := fmt.Errorf("outer: %w", mid)
__check(fmt.Sprint(errors.Unwrap(outer).Error()), "mid: base") }
