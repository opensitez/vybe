// vybe-test: go/functions_patterns_extra/function_composition_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func compose(a func(int) int, b func(int) int) func(int) int { return func(v int) int { return b(a(v)) } }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { fn := compose(func(v int) int { return v + 1 }, func(v int) int { return v * 2 })
__check(fmt.Sprint(fn(5)), "12")
}
