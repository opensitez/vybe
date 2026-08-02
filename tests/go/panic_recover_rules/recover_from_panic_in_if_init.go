// vybe-test: go/panic_recover_rules/recover_from_panic_in_if_init
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { __check(fmt.Sprint(recover()), "if") }()
if x := 1; x > 0 { panic("if") } }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
