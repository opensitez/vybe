// vybe-test: go/method_sets_pointer_value/pointer_embedded_nil_inner_method_call_panic_guard_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type inner struct { n int }
func (i *inner) peek() int { if i == nil { return -1 }
return i.n }
type outer struct { *inner }
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

func main() { var o outer
__p(fmt.Sprint(o.peek())) 
__check("-1")
}
