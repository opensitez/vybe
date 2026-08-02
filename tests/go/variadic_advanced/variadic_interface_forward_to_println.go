// vybe-test: go/variadic_advanced/variadic_interface_forward_to_println
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func dump(parts ...interface{}) { __check(fmt.Sprint(len(parts)), "2") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { dump(true, false) }
