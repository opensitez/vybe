// vybe-test: go/panic_recover_rules/recover_consumes_panic_once
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { __check(fmt.Sprint(recover()), "once")
__check(fmt.Sprint(recover() == nil), "true") }()
panic("once") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
