// vybe-test: go/sort_stable_search/sort_float64s_large_and_small
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

func main() { f := []float64{1e10, 1e-10, 1.0}
sort.Float64s(f)
__check(fmt.Sprint(f[0] < f[1]), "true")
__check(fmt.Sprint(f[2]), "10000000000") }
