// vybe-test: go/sort_stable_search/search_generic_first_match
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

func main() { a := []int{1, 3, 5, 7}
i := sort.Search(len(a), func(k int) bool { return a[k] >= 1 })
__check(fmt.Sprint(i), "0") }
