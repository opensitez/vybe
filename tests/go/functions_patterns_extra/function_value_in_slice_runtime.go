// vybe-test: go/functions_patterns_extra/function_value_in_slice_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { fns := []func(int) int{func(v int) int { return v + 1 }, func(v int) int { return v + 2 }}
__check(fmt.Sprint(fns[1](5)), "7")
}
