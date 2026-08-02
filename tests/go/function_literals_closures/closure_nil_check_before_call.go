// vybe-test: go/function_literals_closures/closure_nil_check_before_call
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var fn func() = nil
if fn != nil { fn() } else { __check(fmt.Sprint("nil"), "nil") } }
