// vybe-test: go/interface_embedding_methods/composite_interface_as_function_arg_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type speaker interface { speak() string }
type louder interface { speaker }
type dog struct{}
func (dog) speak() string { return "woof" }
func echo(value louder) string { return value.speak() }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(echo(dog{})), "woof") }
