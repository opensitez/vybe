// vybe-test: go/interface_nil_comparable/generic_comparable_int_equality
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func equal[T comparable](left T, right T) bool { return left == right }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(equal(3, 3)), "true")
__check(fmt.Sprint(equal(3, 4)), "false") }
