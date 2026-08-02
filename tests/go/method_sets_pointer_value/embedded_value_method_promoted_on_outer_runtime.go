// vybe-test: go/method_sets_pointer_value/embedded_value_method_promoted_on_outer_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type base struct { id int }
func (b base) idVal() int { return b.id }
type shell struct { base }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := shell{base: base{id: 7}}
__check(fmt.Sprint(s.idVal()), "7") }
