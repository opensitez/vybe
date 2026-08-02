// vybe-test: go/functions_patterns_extra/callback_applies_operation_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func apply(v int, fn func(int) int) int { return fn(v) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { result := apply(4, func(v int) int { return v * v })
__check(fmt.Sprint(result), "16")
}
