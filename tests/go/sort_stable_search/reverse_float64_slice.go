// vybe-test: go/sort_stable_search/reverse_float64_slice
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

func main() { f := sort.Float64Slice{1.1, 2.2, 3.3}
sort.Sort(sort.Reverse(f))
__check(fmt.Sprint(f[0]), "3.3")
__check(fmt.Sprint(f[2]), "1.1") }
