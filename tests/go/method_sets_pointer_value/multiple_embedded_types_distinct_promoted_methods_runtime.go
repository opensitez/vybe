// vybe-test: go/method_sets_pointer_value/multiple_embedded_types_distinct_promoted_methods_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type north struct{}
func (north) dir() string { return "N" }
type east struct{}
func (east) dir() string { return "E" }
type compass struct { north
east }
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

func main() { c := compass{}
__p(fmt.Sprint(c.north.dir()))
__p(fmt.Sprint(c.east.dir())) 
__check("N\nE")
}
