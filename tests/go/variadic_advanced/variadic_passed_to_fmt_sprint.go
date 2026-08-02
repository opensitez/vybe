// vybe-test: go/variadic_advanced/variadic_passed_to_fmt_sprint
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func show(parts ...interface{}) string { return fmt.Sprint(parts...) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(show("x", 1)), "x1") }
