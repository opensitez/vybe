// vybe-test: go/panic_recover_rules/recover_after_defer_modifies_named_return
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() (n int) { defer func() { recover()
n = 9 }()
panic("p")
return 1 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(run()), "9") }
