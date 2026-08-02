// vybe-test: go/generics_constraints_extended/generic_ordered_sort_strings
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
import "slices"
func SortStrings[T ~string](s []T) { slices.Sort(s) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { names := []string{"go", "vybe", "lang"}
SortStrings(names)
__check(fmt.Sprint(names[0]), "go") }
