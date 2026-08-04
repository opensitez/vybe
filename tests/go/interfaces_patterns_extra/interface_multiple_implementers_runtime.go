// vybe-test: go/interfaces_patterns_extra/interface_multiple_implementers_runtime
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs

package main
import "fmt"
type speaker interface { speak() string }
type dog struct{}
type cat struct{}
func (dog) speak() string { return "woof" }
func (cat) speak() string { return "meow" }
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

func main() { values := []speaker{dog{}, cat{}}
__p(fmt.Sprint(values[0].speak()))
__p(fmt.Sprint(values[1].speak()))
__check("woof\nmeow")
}
