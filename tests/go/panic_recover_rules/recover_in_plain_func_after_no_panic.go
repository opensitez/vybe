// vybe-test: go/panic_recover_rules/recover_in_plain_func_after_no_panic
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func probe() { __check(fmt.Sprint(recover() == nil), "true") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { probe() }
