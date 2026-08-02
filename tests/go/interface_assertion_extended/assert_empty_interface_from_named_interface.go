// vybe-test: go/interface_assertion_extended/assert_empty_interface_from_named_interface
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type speaker interface { talk() string }
type bot struct{}
func (bot) talk() string { return "hi" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var s speaker = bot{}
var v interface{} = s
b, ok := v.(bot)
__check(fmt.Sprint(b.talk()), "hi")
__check(fmt.Sprint(ok), "true") }
