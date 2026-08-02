// vybe-test: go/defer_lifo_extended/defer_before_return_in_nested_func
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func inner() int { defer __check(fmt.Sprint("in"), "in")
return 1 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer __check(fmt.Sprint("out"), "1")
__check(fmt.Sprint(inner()), "out") }
