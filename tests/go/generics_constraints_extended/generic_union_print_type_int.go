// vybe-test: go/generics_constraints_extended/generic_union_print_type_int
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func Tag[T int | string](v T) string { switch any(v).(type) { case int: return "int"
default: return "string" } }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(Tag(1)), "int") }
