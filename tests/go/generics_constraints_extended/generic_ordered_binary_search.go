// vybe-test: go/generics_constraints_extended/generic_ordered_binary_search
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
import "cmp"
import "slices"
func Find[T cmp.Ordered](s []T, target T) (int, bool) { return slices.BinarySearch(s, target) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { i, ok := Find([]int{1, 3, 5}, 3)
__check(fmt.Sprint(i), "1")
__check(fmt.Sprint(ok), "true") }
