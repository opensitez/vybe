// vybe-test: go/defer_lifo_extended/defer_runs_after_panic_in_same_func_with_recover_outer
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func run() { defer __check(fmt.Sprint("a"), "a")
panic("p") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer func() { recover() }()
run() }
