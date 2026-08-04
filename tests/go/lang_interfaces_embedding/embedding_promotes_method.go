// vybe-test: go/lang_interfaces_embedding/embedding_promotes_method
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs

package main
import "fmt"
type A struct{}
func (A) Hi() string { return "hi" }
type B struct { A }
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

func main() { var b B
__p(fmt.Sprint(b.Hi())) 
__check("hi")
}
