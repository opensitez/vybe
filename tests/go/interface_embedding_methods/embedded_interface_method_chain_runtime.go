// vybe-test: go/interface_embedding_methods/embedded_interface_method_chain_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type builder interface { build() string }
type factory interface { builder }
type widget struct{}
func (widget) build() string { return "built" }
func makeFactory() factory { return widget{} }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(makeFactory().build()), "built") }
