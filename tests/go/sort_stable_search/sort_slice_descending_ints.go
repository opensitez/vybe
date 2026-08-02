// vybe-test: go/sort_stable_search/sort_slice_descending_ints
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

func main() { a := []int{1, 5, 3, 2, 4}
sort.Slice(a, func(i, j int) bool { return a[i] > a[j] })
__check(fmt.Sprint(a[0]), "5")
__check(fmt.Sprint(a[4]), "1") }
