// vybe-test: go/slices_sort_equal_extended/slices_sort_float64s_ascending
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

func main() { s := []float64{3.3, 1.1, 2.2}
slices.Sort(s)
__check(fmt.Sprint(s[0]), "1.1")
__check(fmt.Sprint(s[2]), "3.3") }
