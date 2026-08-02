// vybe-test: go/switch_fallthrough_extended/fallthrough_chain_three_cases
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { switch 1 { case 1: __check(fmt.Sprint(1), "1")
fallthrough
case 2: __check(fmt.Sprint(2), "2")
fallthrough
case 3: __check(fmt.Sprint(3), "3") } }
