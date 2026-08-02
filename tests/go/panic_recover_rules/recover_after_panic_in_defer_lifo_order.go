// vybe-test: go/panic_recover_rules/recover_after_panic_in_defer_lifo_order
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { __check(fmt.Sprint("a"), "b")
recover() }()
defer func() { __check(fmt.Sprint("b"), "a")
panic("p") }() }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
