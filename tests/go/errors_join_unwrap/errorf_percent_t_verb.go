// vybe-test: go/errors_join_unwrap/errorf_percent_t_verb
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { err := fmt.Errorf("type %T", 42)
__check(fmt.Sprint(err.Error()), "type int") }
