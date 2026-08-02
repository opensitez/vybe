// vybe-test: go/sort_slice_find/sort_slice_is_sorted_false
// origin: languages/go/tests/go/test_sort_slice_find.rs

package main
import "fmt"
import "sort"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []int{2,1}
__check(fmt.Sprint(sort.SliceIsSorted(s, func(i,j int) bool { return s[i] < s[j] })), "false") }
