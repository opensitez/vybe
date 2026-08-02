// vybe-test: go/lang_generics_semantics/ordered_constraint_min
// origin: languages/go/tests/go/test_lang_generics_semantics.rs

package main
import "fmt"
import "cmp"
func Smallest[T cmp.Ordered](a, b T) T { if cmp.Less(a,b) { return a }
return b }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Smallest(3,9)), "3") }
