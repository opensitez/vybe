// vybe-test: go/interface_nil_comparable/named_interface_concrete_value_equality
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
type counter interface { count() int }
type tally struct { n int }
func (t tally) count() int { return t.n }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var left counter = tally{n: 5}
var right counter = tally{n: 5}
__check(fmt.Sprint(left == right), "true") }
