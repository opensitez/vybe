// vybe-test: go/type_aliases/defined_to_underlying_explicit_cast
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

func main() { value := Score(13)
__check(fmt.Sprint(int(value)), "13") }
