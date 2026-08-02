// vybe-test: go/defer_panic_recover_extra/defer_order_with_named_functions_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func one() { __check(fmt.Sprint(1), "2") }
func two() { __check(fmt.Sprint(2), "1") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer one()
defer two()
}
