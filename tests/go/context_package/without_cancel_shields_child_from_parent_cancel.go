// vybe-test: go/context_package/without_cancel_shields_child_from_parent_cancel
// origin: languages/go/tests/go/test_context_package.rs

package main
import "fmt"
import "context"
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

func main() { parent, parentCancel := context.WithCancel(context.Background())
child := context.WithoutCancel(parent)
parentCancel()
__p(fmt.Sprint(child.Err() == nil)) 
__check("true")
}
