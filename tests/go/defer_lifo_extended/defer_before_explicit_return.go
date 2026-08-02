// vybe-test: go/defer_lifo_extended/defer_before_explicit_return
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func work() int { defer __check(fmt.Sprint("defer"), "body")
__check(fmt.Sprint("body"), "defer")
return 1 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(work()), "1") }
