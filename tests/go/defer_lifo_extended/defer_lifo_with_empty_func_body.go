// vybe-test: go/defer_lifo_extended/defer_lifo_with_empty_func_body
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func noop() {}
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer noop()
defer __check(fmt.Sprint("x"), "x")
}
