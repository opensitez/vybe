// vybe-test: go/panic_recover_rules/recover_in_deferred_named_func
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func save() { __check(fmt.Sprint(recover()), "named") }
func run() { defer save()
panic("named") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { run() }
