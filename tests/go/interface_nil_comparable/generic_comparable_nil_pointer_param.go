// vybe-test: go/interface_nil_comparable/generic_comparable_nil_pointer_param
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func isNil[T comparable](value T) bool { var zero T
return value == zero }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var p *int
__check(fmt.Sprint(isNil(p)), "true") }
