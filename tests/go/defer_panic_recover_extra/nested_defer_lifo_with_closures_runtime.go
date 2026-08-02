// vybe-test: go/defer_panic_recover_extra/nested_defer_lifo_with_closures_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer func() { __check(fmt.Sprint("first"), "second") }()
defer func() { __check(fmt.Sprint("second"), "first") }()
}
