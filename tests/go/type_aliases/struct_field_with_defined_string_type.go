// vybe-test: go/type_aliases/struct_field_with_defined_string_type
// origin: languages/go/tests/go/test_type_aliases.rs

package main
import "fmt"
type Label string
type item struct { name Label }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := item{name: "vybe"}
__check(fmt.Sprint(value.name), "vybe") }
