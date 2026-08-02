// vybe-test: go/function_types_advanced/return_curried_multiplier_applied_twice
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
func scale(factor int) func(int) int { return func(v int) int { return v * factor } }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { double := scale(2)
triple := scale(3)
__check(fmt.Sprint(double(4)), "8")
__check(fmt.Sprint(triple(4)), "12") }
