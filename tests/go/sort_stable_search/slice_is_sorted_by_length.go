// vybe-test: go/sort_stable_search/slice_is_sorted_by_length
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

func main() { s := []string{"a", "bb", "ccc"}
__check(fmt.Sprint(sort.SliceIsSorted(s, func(i, j int) bool { return len(s[i]) < len(s[j]) })), "true") }
