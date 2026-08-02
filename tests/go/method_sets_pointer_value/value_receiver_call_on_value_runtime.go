// vybe-test: go/method_sets_pointer_value/value_receiver_call_on_value_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type score struct { pts int }
func (s score) total() int { return s.pts }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { v := score{pts: 11}
__check(fmt.Sprint(v.total()), "11") }
