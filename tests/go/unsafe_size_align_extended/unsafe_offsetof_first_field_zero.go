// vybe-test: go/unsafe_size_align_extended/unsafe_offsetof_first_field_zero
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs

package main
import "fmt"
import "unsafe"
type S struct { a int
b int }
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

func main() { __p(fmt.Sprint(unsafe.Offsetof(S{}.a))) 
__check("0")
}
