// vybe-test: go/slices_sort_equal_extended/slices_sort_func_by_absolute_value
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

func main() { s := []int{-3, 1, -2, 4}
slices.SortFunc(s, func(a, b int) int { aa, bb := a, b; if aa < 0 { aa = -aa }; if bb < 0 { bb = -bb }; if aa < bb { return -1 }; if aa > bb { return 1 }; return 0 })
__check(fmt.Sprint(s[0]), "1")
__check(fmt.Sprint(s[3]), "4") }
