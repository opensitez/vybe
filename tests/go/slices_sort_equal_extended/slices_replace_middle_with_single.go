// vybe-test: go/slices_sort_equal_extended/slices_replace_middle_with_single
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

func main() { s := []int{1, 2, 3, 4, 5}
t := slices.Replace(s, 1, 4, 99)
__check(fmt.Sprint(len(t)), "3")
__check(fmt.Sprint(t[1]), "99")
__check(fmt.Sprint(t[2]), "5") }
