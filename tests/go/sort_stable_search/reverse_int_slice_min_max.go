// vybe-test: go/sort_stable_search/reverse_int_slice_min_max
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

func main() { a := sort.IntSlice{1, 2, 3, 4, 5}
sort.Sort(sort.Reverse(a))
__check(fmt.Sprint(a[0]), "5")
__check(fmt.Sprint(a[4]), "1") }
