// vybe-test: go/switch_fallthrough_extended/switch_string_with_fallthrough_to_default
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { switch "a" { case "a": __check(fmt.Sprint("a"), "a")
fallthrough
default: __check(fmt.Sprint("d"), "d") } }
