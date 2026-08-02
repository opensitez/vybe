// vybe-test: go/sort_stable_search/sort_slice_float64_abs
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

func main() { f := []float64{-3.0, 1.0, -2.0, 4.0}
sort.Slice(f, func(i, j int) bool { ai, aj := f[i], f[j]; if ai < 0 { ai = -ai }; if aj < 0 { aj = -aj }; return ai < aj })
__check(fmt.Sprint(f[0]), "1")
__check(fmt.Sprint(f[3]), "-3") }
