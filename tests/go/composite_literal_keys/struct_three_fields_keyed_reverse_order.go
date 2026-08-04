// vybe-test: go/composite_literal_keys/struct_three_fields_keyed_reverse_order
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type record struct { a int
b int
c int }
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

func main() { r := record{c: 3, b: 2, a: 1}
__p(fmt.Sprint(r.a))
__p(fmt.Sprint(r.c))
__check("1\n3")
}
