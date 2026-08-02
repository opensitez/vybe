// vybe-test: go/interface_embedding_methods/promoted_method_returns_string_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type namer interface { name() string }
type labeled interface { namer }
type widget struct{}
func (widget) name() string { return "vybe" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var value labeled = widget{}
__check(fmt.Sprint(value.name()), "vybe") }
