// vybe-test: go/generics_types/generic_matrix_nested_slice
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Matrix[T any] struct { Rows [][]T }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { m := Matrix[int]{Rows: [][]int{{1, 2}, {3, 4}}}
__check(fmt.Sprint(len(m.Rows)), "2")
__check(fmt.Sprint(m.Rows[1][0]), "3") }
