// vybe-test: go/method_sets_pointer_value/value_method_on_nonaddressable_temp_via_interface_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type speaker interface { say() string }
type bot struct{}
func (b bot) say() string { return "beep" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var s speaker = bot{}
__check(fmt.Sprint(s.say()), "beep") }
