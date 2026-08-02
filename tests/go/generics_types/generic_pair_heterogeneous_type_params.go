// vybe-test: go/generics_types/generic_pair_heterogeneous_type_params
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Pair[A, B any] struct { First A
Second B }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { p := Pair[int, string]{First: 9, Second: "go"}
__check(fmt.Sprint(p.First), "9")
__check(fmt.Sprint(p.Second), "go") }
