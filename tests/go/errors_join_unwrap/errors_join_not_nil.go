// vybe-test: go/errors_join_unwrap/errors_join_not_nil
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

func main() { err := errors.Join(errors.New("fail"))
__check(fmt.Sprint(err != nil), "true") }
