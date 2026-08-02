// vybe-test: go/function_types_advanced/func_with_params_cleared_to_nil
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var fn func(int, int) int
fn = func(a int, b int) int { return a + b }
__check(fmt.Sprint(fn(2, 3)), "5")
fn = nil
__check(fmt.Sprint(fn == nil), "true") }
