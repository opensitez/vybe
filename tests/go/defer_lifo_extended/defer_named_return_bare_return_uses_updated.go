// vybe-test: go/defer_lifo_extended/defer_named_return_bare_return_uses_updated
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func work() (n int) { defer func() { n = n + 1 }()
n = 4
return }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(work()), "5") }
