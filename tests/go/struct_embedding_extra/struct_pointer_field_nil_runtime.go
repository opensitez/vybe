// vybe-test: go/struct_embedding_extra/struct_pointer_field_nil_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type node struct { next *node }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := node{}
__check(fmt.Sprint(value.next == nil), "true")
}
