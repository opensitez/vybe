// vybe-test: go/generics_types/generic_triple_three_type_parameters
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Triple[A, B, C any] struct { A A
B B
C C }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { t := Triple[int, string, bool]{A: 1, B: "x", C: true}
__check(fmt.Sprint(t.A), "1")
__check(fmt.Sprint(t.B), "x")
__check(fmt.Sprint(t.C), "true") }
