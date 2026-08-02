// vybe-test: go/slices_sort_equal_extended/slices_sort_stable_func_stable_on_equal_keys
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

func main() { s := []string{"bb", "aa", "ab"}
slices.SortStableFunc(s, func(a, b string) int { la, lb := len(a), len(b); if la < lb { return -1 }; if la > lb { return 1 }; return 0 })
__check(fmt.Sprint(s[0]), "bb")
__check(fmt.Sprint(s[2]), "ab") }
