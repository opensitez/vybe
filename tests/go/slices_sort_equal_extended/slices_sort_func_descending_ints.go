// vybe-test: go/slices_sort_equal_extended/slices_sort_func_descending_ints
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

func main() { s := []int{1, 3, 2}
slices.SortFunc(s, func(a, b int) int { if a > b { return -1 }; if a < b { return 1 }; return 0 })
__check(fmt.Sprint(s[0]), "3")
__check(fmt.Sprint(s[2]), "1") }
