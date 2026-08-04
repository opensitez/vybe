// vybe-test: go/embedding_promotion_extended/four_level_two_step_promotion_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type d struct { n int }
type c struct { d }
type b struct { c }
type a struct { b }
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

func main() { v := a{b: b{c: c{d: d{n: 13}}}}
__p(fmt.Sprint(v.n)) 
__check("13")
}
