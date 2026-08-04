// vybe-test: go/interfaces/method_chain_struct
// origin: languages/go/tests/go/test_interfaces.rs

package main
import "fmt"
type Builder struct { Val int } func (b Builder) Add(n int) Builder { return Builder{Val: b.Val + n} } var __buf string

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

func main() { b := Builder{Val: 0}
b = b.Add(5)
b = b.Add(3)
__p(fmt.Sprint(b.Val))
__check("8")
}
