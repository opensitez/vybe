// vybe-test: go/slices_sort_equal_extended/slices_sort_after_equal_compare
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

func main() { a := []int{3, 1, 2}
b := []int{3, 1, 2}
slices.Sort(a)
slices.Sort(b)
__check(fmt.Sprint(slices.Equal(a, b)), "true")
__check(fmt.Sprint(slices.Compare(a, b)), "0") }
