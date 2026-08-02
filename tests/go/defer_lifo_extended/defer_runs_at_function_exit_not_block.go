// vybe-test: go/defer_lifo_extended/defer_runs_at_function_exit_not_block
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { if true { defer __check(fmt.Sprint("defer"), "block")
__check(fmt.Sprint("block"), "main")
}
__check(fmt.Sprint("main"), "defer") }
