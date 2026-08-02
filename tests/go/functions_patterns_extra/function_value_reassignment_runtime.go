// vybe-test: go/functions_patterns_extra/function_value_reassignment_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func add(a int, b int) int { return a + b }
func mul(a int, b int) int { return a * b }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { op := add
__check(fmt.Sprint(op(2, 3)), "5")
op = mul
__check(fmt.Sprint(op(2, 3)), "6")
}
