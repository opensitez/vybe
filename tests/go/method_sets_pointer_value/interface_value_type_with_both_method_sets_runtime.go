// vybe-test: go/method_sets_pointer_value/interface_value_type_with_both_method_sets_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type speaker interface { say() string }
type bot struct { msg string }
func (b bot) say() string { return b.msg }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var s speaker = bot{msg: "hi"}
__check(fmt.Sprint(s.say()), "hi") }
