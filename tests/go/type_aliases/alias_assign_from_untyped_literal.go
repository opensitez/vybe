// vybe-test: go/type_aliases/alias_assign_from_untyped_literal
// origin: languages/go/tests/go/test_type_aliases.rs

package main
import "fmt"
type Count = int
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var value Count = 7
__check(fmt.Sprint(value), "7") }
