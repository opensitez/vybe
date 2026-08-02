// vybe-test: go/interface_embedding_methods/promoted_interface_method_value_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type greeter interface { greet() string }
type social interface { greeter }
type hi struct{}
func (hi) greet() string { return "hi" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var s social = hi{}
fn := s.greet
__check(fmt.Sprint(fn()), "hi") }
