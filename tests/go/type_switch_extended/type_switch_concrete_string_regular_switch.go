// vybe-test: go/type_switch_extended/type_switch_concrete_string_regular_switch
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
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

func main() { s := "ab"
switch s { case "ab": __p(fmt.Sprint("match"))
default: __p(fmt.Sprint("miss")) } 
__check("match")
}
