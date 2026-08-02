// vybe-test: go/composite_literal_keys/struct_mixed_type_fields_all_keyed
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type item struct { label string
count int
active bool }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { it := item{active: true, label: "vybe", count: 7}
__check(fmt.Sprint(it.label), "vybe")
__check(fmt.Sprint(it.count), "7")
__check(fmt.Sprint(it.active), "true")
}
