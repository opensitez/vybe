// vybe-test: go/defer_lifo_extended/defer_two_funcs_same_name_different_order
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func a() { __check(fmt.Sprint("a"), "b") }
func b() { __check(fmt.Sprint("b"), "a") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer a()
defer b()
}
