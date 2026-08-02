// vybe-test: go/switch_fallthrough_extended/switch_case_list_with_fallthrough
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { switch 2 { case 1, 2: __check(fmt.Sprint("hit"), "hit")
fallthrough
case 3: __check(fmt.Sprint("next"), "next") } }
