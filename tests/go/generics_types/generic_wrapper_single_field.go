// vybe-test: go/generics_types/generic_wrapper_single_field
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Wrapper[T any] struct { V T }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { w := Wrapper[string]{V: "vybe"}
__check(fmt.Sprint(w.V), "vybe") }
