// vybe-test: go/errors_join_unwrap/errorf_percent_v_verb
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { err := fmt.Errorf("data %v", []int{1, 2})
__check(fmt.Sprint(err.Error()), "data [1 2]") }
