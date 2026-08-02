// vybe-test: go/variadic_advanced/mixed_three_fixed_before_spread
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func frame(a string, b string, c string, rest ...string) int { return len(a) + len(b) + len(c) + len(rest) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { tail := []string{"d"}
__check(fmt.Sprint(frame("x", "y", "z", tail...)), "4") }
