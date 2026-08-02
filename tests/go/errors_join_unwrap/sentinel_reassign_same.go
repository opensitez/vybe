// vybe-test: go/errors_join_unwrap/sentinel_reassign_same
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
import "errors"
var ErrOld = errors.New("old")
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ErrOld = errors.New("new")
__check(fmt.Sprint(ErrOld.Error()), "new") }
