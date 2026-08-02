// vybe-test: go/sort_stable_search/sort_ints_already_sorted
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

func main() { a := []int{1, 2, 3, 4}
sort.Ints(a)
__check(fmt.Sprint(a[0]), "1")
__check(fmt.Sprint(a[3]), "4") }
