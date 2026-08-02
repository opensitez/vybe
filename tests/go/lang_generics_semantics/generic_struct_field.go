// vybe-test: go/lang_generics_semantics/generic_struct_field
// origin: languages/go/tests/go/test_lang_generics_semantics.rs

package main
import "fmt"
type Box[T any] struct { V T }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Box[int]{V: 2}.V), "2") }
