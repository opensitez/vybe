// vybe-test: go/function_types_advanced/return_func_selected_by_boolean_flag
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
func pick(positive bool) func(int) int { if positive { return func(v int) int { return v + 1 } }
return func(v int) int { return v - 1 } }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(pick(true)(6)), "7")
__check(fmt.Sprint(pick(false)(6)), "5") }
