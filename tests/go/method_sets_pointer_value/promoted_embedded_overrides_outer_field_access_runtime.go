// vybe-test: go/method_sets_pointer_value/promoted_embedded_overrides_outer_field_access_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type inner struct { x int }
type outer struct { inner
x int }
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

func main() { o := outer{inner: inner{x: 1}, x: 2}
__p(fmt.Sprint(o.x))
__p(fmt.Sprint(o.inner.x)) 
__check("2\n1")
}
