// vybe-test: go/composite_literal_keys/struct_partial_keyed_leaves_zero_values
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type config struct { host string
port int
debug bool }
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

func main() { c := config{port: 8080}
__p(fmt.Sprint(c.port))
__p(fmt.Sprint(c.debug))
__check("8080\nfalse")
}
