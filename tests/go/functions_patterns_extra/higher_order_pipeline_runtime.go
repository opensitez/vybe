// vybe-test: go/functions_patterns_extra/higher_order_pipeline_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func pipe(v int, a func(int) int, b func(int) int) int { return b(a(v)) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(pipe(3, func(v int) int { return v + 1 }, func(v int) int { return v * 2 })), "8")
}
