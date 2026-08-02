// vybe-test: go/slices_sort_equal_extended/slices_equal_different_elements
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

func main() { __check(fmt.Sprint(slices.Equal([]int{1, 2, 3}, []int{1, 9, 3})), "false") }
