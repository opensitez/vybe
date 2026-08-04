// vybe-test: go/composite_literal_keys/struct_keyed_embedded_inner_by_name
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type inner struct { value int }
type outer struct { inner }
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

func main() { o := outer{inner: inner{value: 42}}
__p(fmt.Sprint(o.value))
__check("42")
}
