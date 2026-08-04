// vybe-test: go/interfaces_patterns_extra/interface_value_in_map_runtime
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs

package main
import "fmt"
type speaker interface { speak() string }
type dog struct{}
func (dog) speak() string { return "woof" }
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

func main() { values := map[string]speaker{"pet": dog{}}
__p(fmt.Sprint(values["pet"].speak()))
__check("woof")
}
