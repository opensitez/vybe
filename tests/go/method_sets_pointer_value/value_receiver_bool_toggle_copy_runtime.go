// vybe-test: go/method_sets_pointer_value/value_receiver_bool_toggle_copy_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type flag struct { on bool }
func (f flag) isOn() bool { return f.on }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { f := flag{on: true}
__check(fmt.Sprint(f.isOn()), "true") }
