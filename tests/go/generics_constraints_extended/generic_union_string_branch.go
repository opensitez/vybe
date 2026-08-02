// vybe-test: go/generics_constraints_extended/generic_union_string_branch
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func Len[T string | []byte](v T) int { return len(v) }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Len("go")), "2") }
