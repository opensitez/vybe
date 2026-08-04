// vybe-test: go/composite_literal_keys/array_keyed_mixed_with_positional_continuation
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

func main() { a := [6]int{1: 10, 3: 30, 5}
__p(fmt.Sprint(a[1]))
__p(fmt.Sprint(a[3]))
__p(fmt.Sprint(a[4]))
__p(fmt.Sprint(len(a)))
__check("10\n30\n5\n6")
}
