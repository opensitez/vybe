// vybe-test: go/panic_recover_rules/recover_allows_caller_to_continue
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { recover() }()
panic("x") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run()
__check(fmt.Sprint("after"), "after") }
