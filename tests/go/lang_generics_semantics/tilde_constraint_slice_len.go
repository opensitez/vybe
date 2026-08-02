// vybe-test: go/lang_generics_semantics/tilde_constraint_slice_len
// origin: languages/go/tests/go/test_lang_generics_semantics.rs

package main
import "fmt"
func Len[S ~[]E, E any](s S) int { return len(s) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Len([]int{1,2,3})), "3") }
