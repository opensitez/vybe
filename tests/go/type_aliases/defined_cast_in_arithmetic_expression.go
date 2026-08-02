// vybe-test: go/type_aliases/defined_cast_in_arithmetic_expression
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

func main() { var value Score = 16
__check(fmt.Sprint(int(value) + 1), "17") }
