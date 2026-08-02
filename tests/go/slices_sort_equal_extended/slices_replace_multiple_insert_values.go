// vybe-test: go/slices_sort_equal_extended/slices_replace_multiple_insert_values
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

func main() { s := []int{1, 5}
t := slices.Replace(s, 1, 2, 2, 3, 4)
__check(fmt.Sprint(len(t)), "4")
__check(fmt.Sprint(t[2]), "3")
__check(fmt.Sprint(t[3]), "4") }
