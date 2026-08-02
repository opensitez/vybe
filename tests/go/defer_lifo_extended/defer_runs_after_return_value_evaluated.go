// vybe-test: go/defer_lifo_extended/defer_runs_after_return_value_evaluated
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func work() int { defer __check(fmt.Sprint("defer"), "ret")
return func() int { __check(fmt.Sprint("ret"), "defer")
return 2 }() }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(work()), "2") }
