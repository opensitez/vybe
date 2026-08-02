// vybe-test: go/switch_fallthrough_extended/switch_in_func_returns_via_case
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func label(n int) string { switch n { case 1: return "one"
default: return "many" } }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(label(1)), "one")
__check(fmt.Sprint(label(9)), "many") }
