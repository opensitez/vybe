// vybe-test: go/composite_literal_keys/array_keyed_index_zero_explicit
// origin: languages/go/tests/go/test_composite_literal_keys.rs

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

func main() { a := [5]int{0: 11, 4: 44}
__p(fmt.Sprint(a[0]))
__p(fmt.Sprint(a[1]))
__p(fmt.Sprint(a[4]))
__check("11\n0\n44")
}
