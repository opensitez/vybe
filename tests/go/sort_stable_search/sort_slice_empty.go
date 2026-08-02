// vybe-test: go/sort_stable_search/sort_slice_empty
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

func main() { a := []int{}
sort.Slice(a, func(i, j int) bool { return a[i] < a[j] })
__check(fmt.Sprint(len(a)), "0") }
