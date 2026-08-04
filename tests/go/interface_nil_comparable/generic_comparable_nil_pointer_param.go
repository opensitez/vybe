// vybe-test: go/interface_nil_comparable/generic_comparable_nil_pointer_param
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func isNil[T comparable](value T) bool { var zero T
return value == zero }
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { var p *int
__p(fmt.Sprint(isNil(p))) 
__check("true")
}
