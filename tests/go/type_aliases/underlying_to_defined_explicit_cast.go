// vybe-test: go/type_aliases/underlying_to_defined_explicit_cast
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

func main() { __check(fmt.Sprint(Score(14)), "14") }
