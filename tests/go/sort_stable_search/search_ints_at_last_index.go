// vybe-test: go/sort_stable_search/search_ints_at_last_index
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

func main() { a := []int{2, 4, 6, 8}
__check(fmt.Sprint(sort.SearchInts(a, 8)), "3") }
