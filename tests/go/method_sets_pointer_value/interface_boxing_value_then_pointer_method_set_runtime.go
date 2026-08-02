// vybe-test: go/method_sets_pointer_value/interface_boxing_value_then_pointer_method_set_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type grower interface { grow() }
type plant struct { h int }
func (p *plant) grow() { p.h++ }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { p := &plant{h: 1}
var g grower = p
g.grow()
__check(fmt.Sprint(p.h), "2") }
