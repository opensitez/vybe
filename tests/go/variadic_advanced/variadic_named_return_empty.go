// vybe-test: go/variadic_advanced/variadic_named_return_empty
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func size(items ...int) (n int) { n = len(items)
return }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(size()), "0") }
