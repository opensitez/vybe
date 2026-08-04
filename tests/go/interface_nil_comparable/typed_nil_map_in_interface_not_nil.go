// vybe-test: go/interface_nil_comparable/typed_nil_map_in_interface_not_nil
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

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

func main() { var m map[string]int
var value interface{} = m
__p(fmt.Sprint(value == nil)) 
__check("false")
}
