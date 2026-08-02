// vybe-test: go/slices_sort_equal_extended/slices_equal_nil_vs_empty
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

func main() { var a []int
b := []int{}
__check(fmt.Sprint(slices.Equal(a, b)), "true") }
