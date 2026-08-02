// vybe-test: go/panic_recover_rules/defer_recover_in_method_on_struct
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
type safe struct{}
func (s safe) run() { defer func() { __check(fmt.Sprint(recover()), "m") }()
panic("m") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { safe{}.run() }
