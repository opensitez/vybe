// vybe-test: go/method_sets_pointer_value/embedded_pointer_method_promoted_on_outer_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type base struct { n int }
func (b *base) double() { b.n *= 2 }
type shell struct { base }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := shell{base: base{n: 3}}
s.double()
__check(fmt.Sprint(s.n), "6") }
