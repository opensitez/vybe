// vybe-test: go/generics_constraints_extended/generic_union_int_branch
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func Twice[T int | int64](v T) T { return v + v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Twice(5)), "10") }
