// vybe-test: go/sort_stable_search/sort_slice_bool_custom
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

func main() { a := []int{1, 2, 3, 4, 5, 6}
sort.Slice(a, func(i, j int) bool { return a[i]%2 > a[j]%2 })
__check(fmt.Sprint(a[0]%2), "1")
__check(fmt.Sprint(a[5]%2), "0") }
