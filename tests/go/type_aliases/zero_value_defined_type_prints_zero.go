// vybe-test: go/type_aliases/zero_value_defined_type_prints_zero
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

func main() { var value Score
__check(fmt.Sprint(int(value)), "0") }
