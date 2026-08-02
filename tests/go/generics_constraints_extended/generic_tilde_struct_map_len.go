// vybe-test: go/generics_constraints_extended/generic_tilde_struct_map_len
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
type StrMap map[string]int
func Size[M ~map[string]int](m M) int { return len(m) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Size(StrMap{"a": 1, "b": 2})), "2") }
