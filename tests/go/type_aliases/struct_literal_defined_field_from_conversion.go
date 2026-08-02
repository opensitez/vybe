// vybe-test: go/type_aliases/struct_literal_defined_field_from_conversion
// origin: languages/go/tests/go/test_type_aliases.rs

package main
import "fmt"
type Meters int
type segment struct { length Meters }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := segment{length: Meters(19)}
__check(fmt.Sprint(int(value.length)), "19") }
