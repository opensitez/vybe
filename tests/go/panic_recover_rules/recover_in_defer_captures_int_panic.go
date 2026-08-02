// vybe-test: go/panic_recover_rules/recover_in_defer_captures_int_panic
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { __check(fmt.Sprint(recover()), "42") }()
panic(42) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
