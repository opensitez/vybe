// vybe-test: go/interface_assertion_extended/recover_after_assert_panic_allows_continue
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

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

func main() { caught := false
func() { defer func() { if recover() != nil { caught = true } }()
var v interface{} = 1
_ = v.(bool) }()
if caught { __p(fmt.Sprint("ok")) } else { __p(fmt.Sprint("no")) } 
__check("ok")
}
