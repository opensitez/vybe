// vybe-test: go/defer_lifo_extended/defer_inside_deferred_func_runs_first_on_exit
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

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
