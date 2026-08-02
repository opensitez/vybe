// vybe-test: go/switch_fallthrough_extended/switch_case_with_single_statement
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { switch 1 { case 1: __check(fmt.Sprint("one"), "one") } }
