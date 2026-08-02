// vybe-test: go/slices_sort_equal_extended/slices_sort_func_by_string_length
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

func main() { s := []string{"go", "vybe", "a"}
slices.SortFunc(s, func(a, b string) int { if len(a) < len(b) { return -1 }; if len(a) > len(b) { return 1 }; return 0 })
__check(fmt.Sprint(s[0]), "a")
__check(fmt.Sprint(s[2]), "vybe") }
