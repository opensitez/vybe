// vybe-test: go/type_aliases/struct_field_with_int_alias
// origin: languages/go/tests/go/test_type_aliases.rs

package main
import "fmt"
type Count = int
type row struct { total Count }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := row{total: 18}
__check(fmt.Sprint(value.total), "18") }
