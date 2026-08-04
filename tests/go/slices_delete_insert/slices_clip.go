// vybe-test: go/slices_delete_insert/slices_clip
// origin: languages/go/tests/go/test_slices_delete_insert.rs

package main
import "fmt"
import "slices"
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

func main() { s := make([]int, 3, 10)
t := slices.Clip(s)
__p(fmt.Sprint(len(t)) + " " + fmt.Sprint(cap(t))) 
__check("3 3")
}
