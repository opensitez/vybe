// vybe-test: go/slice_copy_clear/copy_into_larger_dst
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

func main() { dst := make([]int, 5)
src := []int{7,8}
n := copy(dst, src)
__p(fmt.Sprint(n))
__p(fmt.Sprint(dst[0]))
__p(fmt.Sprint(dst[4])) 
__check("2\n7\n0")
}
