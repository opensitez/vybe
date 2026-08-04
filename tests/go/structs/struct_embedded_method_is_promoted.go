// vybe-test: go/structs/struct_embedded_method_is_promoted
// origin: languages/go/tests/go/test_structs.rs

package main
import "fmt"
type Base struct { ID int }
func (b Base) Describe() int { return b.ID }
type Child struct { Base
Name string }
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

func main() { c := Child{Base: Base{ID: 9}, Name: "x"}
__p(fmt.Sprint(c.Describe())) 
__check("9")
}
