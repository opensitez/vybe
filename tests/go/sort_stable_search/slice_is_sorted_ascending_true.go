// vybe-test: go/sort_stable_search/slice_is_sorted_ascending_true
// origin: languages/go/tests/go/test_sort_stable_search.rs

package main
import "fmt"
import "sort"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []int{1, 2, 3, 4}
__check(fmt.Sprint(sort.SliceIsSorted(s, func(i, j int) bool { return s[i] < s[j] })), "true") }
