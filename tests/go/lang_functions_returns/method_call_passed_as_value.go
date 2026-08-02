// vybe-test: go/lang_functions_returns/method_call_passed_as_value
// origin: languages/go/tests/go/test_lang_functions_returns.rs

package main
import "fmt"
type S struct{}
func (S) ID() int { return 7 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var s S
__check(fmt.Sprint(s.ID()), "7") }
