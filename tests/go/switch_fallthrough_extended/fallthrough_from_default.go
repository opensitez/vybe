// vybe-test: go/switch_fallthrough_extended/fallthrough_from_default
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { switch 9 { default: __check(fmt.Sprint("d"), "d")
fallthrough
case 9: __check(fmt.Sprint("nine"), "nine") } }
