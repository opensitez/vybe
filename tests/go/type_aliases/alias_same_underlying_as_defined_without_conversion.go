// vybe-test: go/type_aliases/alias_same_underlying_as_defined_without_conversion
// origin: languages/go/tests/go/test_type_aliases.rs

package main
import "fmt"
type Units int
type Reading = Units
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

func main() { var base Units = 4
var view Reading = base
__p(fmt.Sprint(int(view))) 
__check("4")
}
