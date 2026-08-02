// vybe-test: go/slices_sort_equal_extended/slices_binary_search_after_last
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

func main() { i, ok := slices.BinarySearch([]int{2, 4, 6}, 9)
__check(fmt.Sprint(i), "3")
__check(fmt.Sprint(ok), "false") }
