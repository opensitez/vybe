// vybe-test: go/sort_stable_search/search_generic_int_slice
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

func main() { a := []int{10, 20, 30, 40}
i := sort.Search(len(a), func(k int) bool { return a[k] >= 25 })
__check(fmt.Sprint(i), "2")
__check(fmt.Sprint(a[i]), "30") }
