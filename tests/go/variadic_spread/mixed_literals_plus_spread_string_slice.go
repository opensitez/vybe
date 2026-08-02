// vybe-test: go/variadic_spread/mixed_literals_plus_spread_string_slice
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func join3(a string, b string, rest ...string) int { return len(rest) + len(a) + len(b) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { tail := []string{"c", "d"}
__check(fmt.Sprint(join3("x", "y", tail...)), "4")
}
