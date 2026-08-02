// vybe-test: go/interface_embedding_methods/overlapping_method_unified_impl_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type resetterA interface { reset() int }
type resetterB interface { reset() int }
type dualReset interface { resetterA
resetterB }
type engine struct { ticks int }
func (e *engine) reset() int { e.ticks = 0
return e.ticks }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := &engine{ticks: 5}
var d dualReset = value
__check(fmt.Sprint(d.reset()), "0") }
