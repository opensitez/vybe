// vybe-test: go/defer_panic_recover_extra/recover_in_nested_function_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func run() { defer func() { __check(fmt.Sprint(recover() != nil), "true") }()
panic("boom") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
