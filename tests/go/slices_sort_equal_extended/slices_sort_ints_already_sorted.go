// vybe-test: go/slices_sort_equal_extended/slices_sort_ints_already_sorted
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

func main() { s := []int{1, 2, 3}
slices.Sort(s)
__check(fmt.Sprint(slices.IsSorted(s)), "true") }
