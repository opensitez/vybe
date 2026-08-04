// vybe-test: go/type_aliases/map_with_defined_type_value_cast
// origin: languages/go/tests/go/test_type_aliases.rs

package main
import "fmt"
type Level int
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

func main() { values := map[string]Level{"a": Level(20)}
__p(fmt.Sprint(int(values["a"]))) 
__check("20")
}
