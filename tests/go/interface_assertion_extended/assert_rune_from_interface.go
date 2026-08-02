// vybe-test: go/interface_assertion_extended/assert_rune_from_interface
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var v interface{} = rune(8364)
r, ok := v.(rune)
__check(fmt.Sprint(int(r)), "8364")
__check(fmt.Sprint(ok), "true") }
