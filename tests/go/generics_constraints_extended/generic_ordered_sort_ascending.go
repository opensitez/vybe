// vybe-test: go/generics_constraints_extended/generic_ordered_sort_ascending
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
import "cmp"
import "slices"
func SortAsc[T cmp.Ordered](s []T) { slices.Sort(s) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { data := []int{3, 1, 2}
SortAsc(data)
__check(fmt.Sprint(data[0]), "1")
__check(fmt.Sprint(data[2]), "3") }
