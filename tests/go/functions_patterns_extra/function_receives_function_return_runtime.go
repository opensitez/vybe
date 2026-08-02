// vybe-test: go/functions_patterns_extra/function_receives_function_return_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func builder() func(int) int { return func(v int) int { return v + 2 } }
func apply(v int, fn func(int) int) int { return fn(v) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(apply(5, builder())), "7")
}
