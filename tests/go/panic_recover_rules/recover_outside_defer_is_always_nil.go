// vybe-test: go/panic_recover_rules/recover_outside_defer_is_always_nil
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(recover() == nil), "true")
__check(fmt.Sprint(recover() == nil), "true") }
