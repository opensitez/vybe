// vybe-test: go/type_aliases/defined_assign_from_untyped_literal
// origin: languages/go/tests/go/test_type_aliases.rs

package main
import "fmt"
type Score int
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var value Score = 12
__check(fmt.Sprint(value), "12") }
