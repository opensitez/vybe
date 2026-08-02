// vybe-test: go/defer_panic_recover_extra/defer_updates_outer_variable_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { total := 1
func() { defer func() { total = 8 }() }()
__check(fmt.Sprint(total), "8")
}
