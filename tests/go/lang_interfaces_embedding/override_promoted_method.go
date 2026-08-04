// vybe-test: go/lang_interfaces_embedding/override_promoted_method
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs

package main
import "fmt"
type A struct{}
func (A) Name() string { return "A" }
type B struct { A }
func (B) Name() string { return "B" }
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

func main() { __p(fmt.Sprint(B{}.Name())) 
__check("B")
}
