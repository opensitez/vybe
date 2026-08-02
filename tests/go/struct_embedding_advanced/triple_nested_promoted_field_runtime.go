// vybe-test: go/struct_embedding_advanced/triple_nested_promoted_field_runtime
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs

package main
import "fmt"
type leaf struct { value int }
type branch struct { leaf }
type trunk struct { branch }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := trunk{branch: branch{leaf: leaf{value: 7}}}
__check(fmt.Sprint(value.value), "7")
}
