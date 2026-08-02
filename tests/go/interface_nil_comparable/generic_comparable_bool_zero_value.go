// vybe-test: go/interface_nil_comparable/generic_comparable_bool_zero_value
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func isZero[T comparable](value T) bool { var zero T
return value == zero }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var flag bool
__check(fmt.Sprint(isZero(flag)), "true") }
