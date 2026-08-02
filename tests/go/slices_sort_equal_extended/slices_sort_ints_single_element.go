// vybe-test: go/slices_sort_equal_extended/slices_sort_ints_single_element
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

func main() { s := []int{7}
slices.Sort(s)
__check(fmt.Sprint(s[0]), "7") }
