// vybe-test: go/slices_sort_equal_extended/slices_replace_end_range
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

func main() { s := []int{1, 2, 3}
t := slices.Replace(s, 2, 3, 9)
__check(fmt.Sprint(t[1]), "2")
__check(fmt.Sprint(t[2]), "9") }
