// vybe-test: go/slices_sort_equal_extended/slices_is_sorted_func_custom_order
// origin: languages/go/tests/go/test_slices_sort_equal_extended.rs

package main
import "fmt"
import "slices"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []int{3, 2, 1}
__check(fmt.Sprint(slices.IsSortedFunc(s, func(a, b int) int { if a > b { return -1 }; if a < b { return 1 }; return 0 })), "true") }
