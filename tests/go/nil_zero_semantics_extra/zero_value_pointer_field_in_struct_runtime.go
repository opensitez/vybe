// vybe-test: go/nil_zero_semantics_extra/zero_value_pointer_field_in_struct_runtime
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs

package main
import "fmt"
type node struct { next *node }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var n node
__check(fmt.Sprint(n.next == nil), "true")
}
