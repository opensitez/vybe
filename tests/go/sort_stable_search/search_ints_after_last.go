// vybe-test: go/sort_stable_search/search_ints_after_last
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

func main() { a := []int{10, 20, 30}
__check(fmt.Sprint(sort.SearchInts(a, 99)), "3") }
