// vybe-test: go/composite_literal_keys/struct_keyed_pointer_field_inline
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type node struct { value int }
type holder struct { head *node }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { h := holder{head: &node{value: 17}}
__check(fmt.Sprint(h.head.value), "17")
}
