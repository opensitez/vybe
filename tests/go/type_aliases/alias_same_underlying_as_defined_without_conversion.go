// vybe-test: go/type_aliases/alias_same_underlying_as_defined_without_conversion
// origin: languages/go/tests/go/test_type_aliases.rs

package main
import "fmt"
type Units int
type Reading = Units
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var base Units = 4
var view Reading = base
__check(fmt.Sprint(int(view)), "4") }
