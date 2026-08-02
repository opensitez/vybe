// vybe-test: go/variadic_advanced/variadic_empty_call_zero_len
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func count(nums ...int) int { return len(nums) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(count()), "0") }
