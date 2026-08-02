// vybe-test: go/defer_panic_variants/defer_registered_inside_deferred_func_runs_first
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer func() { defer __check(fmt.Sprint("inner"), "outer")
__check(fmt.Sprint("outer"), "inner") }()
}
