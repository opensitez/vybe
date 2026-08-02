// vybe-test: go/slices_sort_equal_extended/slices_index_func_no_match
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

func main() { __check(fmt.Sprint(slices.IndexFunc([]int{1, 3, 5}, func(v int) bool { return v > 10 })), "-1") }
