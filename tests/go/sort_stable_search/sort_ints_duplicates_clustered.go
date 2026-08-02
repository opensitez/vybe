// vybe-test: go/sort_stable_search/sort_ints_duplicates_clustered
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

func main() { a := []int{2, 1, 2, 1, 3}
sort.Ints(a)
__check(fmt.Sprint(a[0]), "1")
__check(fmt.Sprint(a[4]), "3") }
