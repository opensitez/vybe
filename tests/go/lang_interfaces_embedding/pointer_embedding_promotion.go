// vybe-test: go/lang_interfaces_embedding/pointer_embedding_promotion
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs

package main
import "fmt"
type A struct{}
func (A) X() int { return 1 }
type B struct { *A }
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

func main() { b := B{A: &A{}}
__p(fmt.Sprint(b.X())) 
__check("1")
}
