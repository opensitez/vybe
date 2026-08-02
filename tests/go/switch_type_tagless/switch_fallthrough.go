// vybe-test: go/switch_type_tagless/switch_fallthrough
// origin: languages/go/tests/go/test_switch_type_tagless.rs

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
