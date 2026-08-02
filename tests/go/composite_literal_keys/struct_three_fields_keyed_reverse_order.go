// vybe-test: go/composite_literal_keys/struct_three_fields_keyed_reverse_order
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type record struct { a int
b int
c int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { r := record{c: 3, b: 2, a: 1}
__check(fmt.Sprint(r.a), "1")
__check(fmt.Sprint(r.c), "3")
}
