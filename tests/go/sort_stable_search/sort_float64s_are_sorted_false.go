// vybe-test: go/sort_stable_search/sort_float64s_are_sorted_false
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

func main() { __check(fmt.Sprint(sort.Float64sAreSorted([]float64{2.0, 1.0})), "false") }
