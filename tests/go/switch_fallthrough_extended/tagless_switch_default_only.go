// vybe-test: go/switch_fallthrough_extended/tagless_switch_default_only
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { switch { default: __check(fmt.Sprint("def"), "def") } }
