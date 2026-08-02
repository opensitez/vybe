// vybe-test: go/type_aliases/slice_of_defined_type_element_cast
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

func main() { values := []Score{Score(1), Score(2)}
__check(fmt.Sprint(int(values[1])), "2") }
