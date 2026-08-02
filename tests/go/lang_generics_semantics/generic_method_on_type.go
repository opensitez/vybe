// vybe-test: go/lang_generics_semantics/generic_method_on_type
// origin: languages/go/tests/go/test_lang_generics_semantics.rs

package main
import "fmt"
type S[T comparable] struct { V T }
func (s S[T]) Same(u T) bool { return s.V == u }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(S[int]{1}.Same(1)), "true") }
