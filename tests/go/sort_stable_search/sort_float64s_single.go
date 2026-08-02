// vybe-test: go/sort_stable_search/sort_float64s_single
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

func main() { f := []float64{3.14}
sort.Float64s(f)
__check(fmt.Sprint(f[0]), "3.14") }
