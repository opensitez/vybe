// vybe-test: go/sort_stable_search/search_float64s_middle
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

func main() { f := []float64{1.0, 2.5, 4.0}
__check(fmt.Sprint(sort.SearchFloat64s(f, 2.5)), "1") }
