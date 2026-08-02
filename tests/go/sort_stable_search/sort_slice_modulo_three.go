// vybe-test: go/sort_stable_search/sort_slice_modulo_three
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

func main() { a := []int{7, 2, 5, 8, 1}
sort.Slice(a, func(i, j int) bool { return a[i]%3 < a[j]%3 })
__check(fmt.Sprint(a[0]%3), "1")
__check(fmt.Sprint(a[4]%3), "2") }
