// vybe-test: go/struct_embedding_advanced/outer_method_shadows_embedded_runtime
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs

package main
import "fmt"
type inner struct{}
func (inner) label() string { return "inner" }
type outer struct { inner }
func (outer) label() string { return "outer" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := outer{}
__check(fmt.Sprint(value.label()), "outer")
}
