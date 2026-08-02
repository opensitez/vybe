// vybe-test: go/variadic_advanced/mixed_one_fixed_plus_variadic
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func prefix(tag string, msgs ...string) int { return len(tag) + len(msgs) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(prefix("ERR", "a", "b")), "5") }
