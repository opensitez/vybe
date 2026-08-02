// vybe-test: go/method_sets_pointer_value/interface_from_literal_pointer_with_pointer_method_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type resetter interface { reset() }
type timer struct { ticks int }
func (t *timer) reset() { t.ticks = 0 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var r resetter = &timer{ticks: 5}
r.reset()
__check(fmt.Sprint(r.(*timer).ticks), "0") }
