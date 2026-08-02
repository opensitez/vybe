// vybe-test: go/function_literals_closures/return_closure_captures_parameter
// origin: languages/go/tests/go/test_function_literals_closures.rs

package main
import "fmt"
func scale(factor int) func(int) int { return func(v int) int { return v * factor } }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { triple := scale(3)
__check(fmt.Sprint(triple(4)), "12") }
