// vybe-test: go/interface_embedding_methods/triple_embedded_interface_dispatch_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type leaf interface { tag() string }
type branch interface { leaf }
type trunk interface { branch }
type node struct{}
func (node) tag() string { return "deep" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var t trunk = node{}
__check(fmt.Sprint(t.tag()), "deep") }
