// vybe-test: go/slice_copy_clear/append_grows_len_and_maybe_cap
// origin: languages/go/tests/go/test_slice_copy_clear.rs

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

func main() { s := make([]int, 0, 2)
s = append(s, 1, 2, 3)
__p(fmt.Sprint(len(s)))
__p(fmt.Sprint(cap(s) >= 3)) 
__check("3\ntrue")
}
