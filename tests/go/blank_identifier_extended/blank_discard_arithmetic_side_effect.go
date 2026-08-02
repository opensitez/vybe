// vybe-test: go/blank_identifier_extended/blank_discard_arithmetic_side_effect
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { x := 5
_ = x + 3
__check(fmt.Sprint(x), "5") }
