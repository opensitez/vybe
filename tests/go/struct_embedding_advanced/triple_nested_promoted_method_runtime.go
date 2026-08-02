// vybe-test: go/struct_embedding_advanced/triple_nested_promoted_method_runtime
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs

package main
import "fmt"
type leaf struct{}
func (leaf) tag() string { return "deep" }
type branch struct { leaf }
type trunk struct { branch }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := trunk{}
__check(fmt.Sprint(value.tag()), "deep")
}
