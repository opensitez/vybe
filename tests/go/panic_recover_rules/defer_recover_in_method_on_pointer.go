// vybe-test: go/panic_recover_rules/defer_recover_in_method_on_pointer
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
type safe struct{}
func (s *safe) run() { defer func() { __p(fmt.Sprint(recover() != nil)) }()
panic(1) }
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

func main() { (&safe{}).run() 
__check("true")
}
