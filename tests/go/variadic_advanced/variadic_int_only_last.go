// vybe-test: go/variadic_advanced/variadic_int_only_last
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func last(nums ...int) int { return nums[len(nums)-1] }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(last(3, 7, 11)), "11") }
