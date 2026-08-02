// vybe-test: go/generics_constraints_extended/generic_comparable_not_for_ordered_sort
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func CountMap[K comparable, V any](m map[K]V) int { return len(m) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(CountMap(map[bool]int{true: 1, false: 2})), "2") }
