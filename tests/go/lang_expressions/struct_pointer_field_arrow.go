// vybe-test: go/lang_expressions/struct_pointer_field_arrow
// origin: languages/go/tests/go/test_lang_expressions.rs

package main
import "fmt"
type P struct { N int }
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

func main() { p := &P{N:1}
p.N = 2
__p(fmt.Sprint(p.N)) 
__check("2")
}
