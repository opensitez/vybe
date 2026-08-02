// vybe-test: go/pointer_receivers_advanced/pointer_receiver_bool_field_toggle_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type flag struct { on bool }
func (f *flag) flip() { f.on = !f.on }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := &flag{on: false}
value.flip()
__check(fmt.Sprint(value.on), "true")
}
