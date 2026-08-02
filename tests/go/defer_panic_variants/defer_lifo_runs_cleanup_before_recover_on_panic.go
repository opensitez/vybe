// vybe-test: go/defer_panic_variants/defer_lifo_runs_cleanup_before_recover_on_panic
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func run() { defer __check(fmt.Sprint("cleanup"), "cleanup")
defer func() { recover() }()
panic("stop") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run()
__check(fmt.Sprint("done"), "done") }
