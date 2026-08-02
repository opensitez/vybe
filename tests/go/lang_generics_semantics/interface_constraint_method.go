// vybe-test: go/lang_generics_semantics/interface_constraint_method
// origin: languages/go/tests/go/test_lang_generics_semantics.rs

package main
import "fmt"
type Stringer interface { String() string }
func Print[T Stringer](v T) { __check(fmt.Sprint(v.String()), "m") }
type My int
func (m My) String() string { return "m" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { Print(My(0)) }
