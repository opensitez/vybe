// vybe-test: go/method_sets_pointer_value/pointer_type_satisfies_interface_with_pointer_method_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type writer interface { write(int) }
type pad struct { n int }
func (p *pad) write(v int) { p.n = v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var w writer = &pad{}
w.write(9)
__check(fmt.Sprint(w.(*pad).n), "9") }
