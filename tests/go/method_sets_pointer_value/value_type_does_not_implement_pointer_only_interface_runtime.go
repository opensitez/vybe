// vybe-test: go/method_sets_pointer_value/value_type_does_not_implement_pointer_only_interface_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type mutator interface { set(int) }
type gauge struct { n int }
func (g *gauge) set(v int) { g.n = v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { g := gauge{}
var m mutator = &g
m.set(4)
__check(fmt.Sprint(g.n), "4") }
