// vybe-test: go/defer_lifo_extended/defer_named_return_bool_set_true
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func work() (ok bool) { defer func() { ok = true }()
return false }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(work()), "true") }
