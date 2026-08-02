// vybe-test: go/generics_constraints_extended/generic_any_append_to_slice
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func Append[T any](s []T, vals ...T) []T { return append(s, vals...) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(len(Append([]int{1}, 2, 3))), "3") }
