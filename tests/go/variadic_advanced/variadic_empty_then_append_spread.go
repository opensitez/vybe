// vybe-test: go/variadic_advanced/variadic_empty_then_append_spread
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func lenAfter(base []int, more ...int) int { combined := append(base, more...)
return len(combined) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(lenAfter([]int{1}, 2, 3)), "3") }
