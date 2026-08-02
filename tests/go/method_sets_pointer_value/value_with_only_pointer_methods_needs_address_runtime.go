// vybe-test: go/method_sets_pointer_value/value_with_only_pointer_methods_needs_address_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type latch struct { on bool }
func (l *latch) flip() { l.on = !l.on }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { l := latch{on: false}
l.flip()
__check(fmt.Sprint(l.on), "true") }
