// vybe-test: go/switch_fallthrough_extended/fallthrough_skips_case_check
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { x := 1
switch x { case 1: __check(fmt.Sprint(1), "1")
fallthrough
case 3: __check(fmt.Sprint(3), "3") } }
