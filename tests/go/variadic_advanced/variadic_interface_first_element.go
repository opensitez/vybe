// vybe-test: go/variadic_advanced/variadic_interface_first_element
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func first(values ...interface{}) interface{} { return values[0] }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(first(99, "x")), "99") }
