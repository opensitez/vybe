// vybe-test: go/slices_sort_equal_extended/slices_binary_search_func_not_found_desc
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

func main() { s := []int{9, 7, 5}
i, ok := slices.BinarySearchFunc(s, 6, func(a, b int) int { if a > b { return -1 }; if a < b { return 1 }; return 0 })
__check(fmt.Sprint(i), "2")
__check(fmt.Sprint(ok), "false") }
