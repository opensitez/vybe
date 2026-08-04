// vybe-test: go/methods/method_same_name_different_types
// origin: languages/go/tests/go/test_methods.rs

package main
import "fmt"
type A struct{}
func (A) Print() { __p(fmt.Sprint("A")) }
type B struct{}
func (B) Print() { __p(fmt.Sprint("B")) }
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

func main() { a := A{}
b := B{}
a.Print()
b.Print()
__check("A\nB")
}
