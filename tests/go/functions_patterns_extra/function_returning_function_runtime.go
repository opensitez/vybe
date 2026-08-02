// vybe-test: go/functions_patterns_extra/function_returning_function_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func maker(step int) func(int) int { return func(v int) int { return v + step } }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { next := maker(3)
__check(fmt.Sprint(next(4)), "7")
}
