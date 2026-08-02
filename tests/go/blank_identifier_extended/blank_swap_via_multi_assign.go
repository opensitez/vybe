// vybe-test: go/blank_identifier_extended/blank_swap_via_multi_assign
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a, b := 1, 2
a, b = b, a
_, _ = a, b
__check(fmt.Sprint(a), "2")
__check(fmt.Sprint(b), "1") }
