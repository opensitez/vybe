// vybe-test: go/panic_recover_rules/panic_in_called_func_recovered_by_caller_defer
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func boom() { panic("deep") }
func run() { defer func() { __check(fmt.Sprint(recover()), "deep") }()
boom() }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
