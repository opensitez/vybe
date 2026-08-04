// vybe-test: go/composite_literal_keys/slice_keyed_mixed_with_positional
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

func main() { s := []int{1: 10, 20, 30}
__p(fmt.Sprint(s[1]))
__p(fmt.Sprint(s[2]))
__p(fmt.Sprint(s[3]))
__p(fmt.Sprint(len(s)))
__check("10\n20\n30\n4")
}
