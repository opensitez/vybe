// vybe-test: go/variadic_spread/forward_variadic_double_delegate_chain
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func end(nums ...int) int { return len(nums) }
func mid(nums ...int) int { return end(nums...) }
func start(nums ...int) int { return mid(nums...) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(start(1, 2, 3, 4)), "4")
}
