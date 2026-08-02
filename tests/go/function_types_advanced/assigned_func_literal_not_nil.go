// vybe-test: go/function_types_advanced/assigned_func_literal_not_nil
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var fn func(int) int
__check(fmt.Sprint(fn == nil), "true")
fn = func(v int) int { return v + 1 }
__check(fmt.Sprint(fn == nil), "false") }
