// vybe-test: go/type_aliases/struct_field_with_slice_alias
// origin: languages/go/tests/go/test_type_aliases.rs

package main
import "fmt"
type IDs = []int
type batch struct { items IDs }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := batch{items: IDs{3, 4}}
__check(fmt.Sprint(len(value.items)), "2")
__check(fmt.Sprint(value.items[1]), "4") }
