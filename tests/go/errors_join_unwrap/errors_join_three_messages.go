// vybe-test: go/errors_join_unwrap/errors_join_three_messages
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

func main() { err := errors.Join(errors.New("one"), errors.New("two"), errors.New("three"))
__check(fmt.Sprint(err.Error()), "one\ntwo\nthree") }
