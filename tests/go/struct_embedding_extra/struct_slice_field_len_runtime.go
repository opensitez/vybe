// vybe-test: go/struct_embedding_extra/struct_slice_field_len_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type bag struct { values []int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := bag{values: []int{2, 4, 6}}
__check(fmt.Sprint(len(value.values)), "3")
}
