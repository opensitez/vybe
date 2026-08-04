// vybe-test: go/composite_literals_extra/slice_literal_of_structs_runtime
// origin: languages/go/tests/go/test_composite_literals_extra.rs

package main
import "fmt"
type user struct { name string }
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

func main() { users := []user{{name: "a"}, {name: "b"}}
__p(fmt.Sprint(users[0].name))
__check("a")
}
