// vybe-test: go/switch_fallthrough_extended/fallthrough_into_default
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { switch 1 { case 1: __check(fmt.Sprint("one"), "one")
fallthrough
default: __check(fmt.Sprint("def"), "def") } }
