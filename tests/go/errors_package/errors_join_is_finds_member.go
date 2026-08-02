// vybe-test: go/errors_package/errors_join_is_finds_member
// origin: languages/go/tests/go/test_errors_package.rs

package main
import "fmt"
import "errors"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { one := errors.New("one")
two := errors.New("two")
joined := errors.Join(one, two)
__check(fmt.Sprint(errors.Is(joined, one)), "true") }
