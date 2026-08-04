// vybe-test: go/context_package/with_cancel_local_cancel_leaves_parent_active
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

func main() { parent, _ := context.WithCancel(context.Background())
child, childCancel := context.WithCancel(parent)
childCancel()
__p(fmt.Sprint(parent.Err() == nil))
__p(fmt.Sprint(child.Err() == context.Canceled)) 
__check("true\ntrue")
}
