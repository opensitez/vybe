// vybe-test: go/sort_stable_search/reverse_then_search
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

func main() { a := sort.IntSlice{1, 3, 5, 7}
sort.Sort(sort.Reverse(a))
__check(fmt.Sprint(a[0]), "7")
__check(fmt.Sprint(sort.SearchInts(a, 5)), "2") }
