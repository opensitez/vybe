// vybe-test: go/defer_lifo_extended/defer_with_variadic_helper
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func show(parts ...string) { __check(fmt.Sprint(len(parts)), "2") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer show("a", "b")
}
