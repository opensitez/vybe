// vybe-test: go/slices_sort_equal_extended/slices_sort_stable_func_preserves_equal_order
// origin: languages/go/tests/go/test_slices_sort_equal_extended.rs

package main
import "fmt"
import "slices"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { type pair struct { k int
ord int }
s := []pair{{k: 1, ord: 0}, {k: 2, ord: 0}, {k: 1, ord: 1}}
slices.SortStableFunc(s, func(a, b pair) int { if a.k < b.k { return -1 }; if a.k > b.k { return 1 }; return 0 })
__check(fmt.Sprint(s[0].ord), "0")
__check(fmt.Sprint(s[1].ord), "1") }
