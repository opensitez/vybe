// vybe-test: go/generics_types/generic_pair_homogeneous_fields
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Pair[T any] struct { First, Second T }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { p := Pair[int]{First: 3, Second: 7}
__check(fmt.Sprint(p.First), "3")
__check(fmt.Sprint(p.Second), "7") }
