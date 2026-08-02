// vybe-test: go/switch_fallthrough_extended/fallthrough_to_next_case
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
switch x { case 1: __check(fmt.Sprint("a"), "a")
fallthrough
case 2: __check(fmt.Sprint("b"), "b") } }
