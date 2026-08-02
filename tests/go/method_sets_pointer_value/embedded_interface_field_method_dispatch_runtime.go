// vybe-test: go/method_sets_pointer_value/embedded_interface_field_method_dispatch_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type speaker interface { talk() string }
type bot struct{}
func (bot) talk() string { return "bot" }
type host struct { speaker }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { h := host{speaker: bot{}}
__check(fmt.Sprint(h.talk()), "bot") }
