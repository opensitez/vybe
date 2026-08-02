// vybe-test: go/method_values/interface_method_value
// origin: languages/go/tests/go/test_method_values.rs

package main
import "fmt"
type greeter interface { greet() string }
type hi struct{}
func (h hi) greet() string { return "hi" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var g greeter = hi{}
f := g.greet
__check(fmt.Sprint(f()), "hi") }
