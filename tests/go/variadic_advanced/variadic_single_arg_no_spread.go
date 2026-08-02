// vybe-test: go/variadic_advanced/variadic_single_arg_no_spread
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func only(n ...int) int { return n[0] }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(only(42)), "42") }
