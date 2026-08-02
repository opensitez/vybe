// vybe-test: go/defer_lifo_extended/defer_multiple_modify_same_named_return
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func work() (n int) { defer func() { n += 1 }()
defer func() { n += 10 }()
return 0 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(work()), "11") }
