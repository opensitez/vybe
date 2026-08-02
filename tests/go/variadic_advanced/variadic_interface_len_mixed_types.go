// vybe-test: go/variadic_advanced/variadic_interface_len_mixed_types
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func pack(values ...interface{}) int { return len(values) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(pack(1, "two", true, 4.0)), "4") }
