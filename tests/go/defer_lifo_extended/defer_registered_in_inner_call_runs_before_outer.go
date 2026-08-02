// vybe-test: go/defer_lifo_extended/defer_registered_in_inner_call_runs_before_outer
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func inner() { defer __check(fmt.Sprint("inner"), "inner")
}
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer __check(fmt.Sprint("outer"), "outer")
inner()
}
