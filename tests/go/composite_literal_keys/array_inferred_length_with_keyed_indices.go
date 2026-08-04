// vybe-test: go/composite_literal_keys/array_inferred_length_with_keyed_indices
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

func main() { a := [...]int{3: 9, 4: 10}
__p(fmt.Sprint(len(a)))
__p(fmt.Sprint(a[3]))
__p(fmt.Sprint(a[0]))
__check("5\n9\n0")
}
