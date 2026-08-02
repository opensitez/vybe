// vybe-test: go/generics_constraints_extended/generic_comparable_array_key
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func LenMap[K comparable, V any](m map[K]V) int { return len(m) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(LenMap(map[[2]int]string{[2]int{1, 2}: "pair"})), "1") }
