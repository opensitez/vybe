// vybe-test: go/variadic_advanced/spread_nil_slice_len_zero
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func size(items ...int) int { return len(items) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var s []int
__check(fmt.Sprint(size(s...)), "0") }
