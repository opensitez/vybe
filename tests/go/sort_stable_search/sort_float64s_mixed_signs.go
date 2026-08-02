// vybe-test: go/sort_stable_search/sort_float64s_mixed_signs
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

func main() { f := []float64{1.1, -1.1, 0.0, 2.2, -2.2}
sort.Float64s(f)
__check(fmt.Sprint(f[0]), "-2.2")
__check(fmt.Sprint(f[4]), "2.2") }
