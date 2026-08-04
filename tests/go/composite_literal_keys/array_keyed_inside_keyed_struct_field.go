// vybe-test: go/composite_literal_keys/array_keyed_inside_keyed_struct_field
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type box struct { data [4]int }
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

func main() { b := box{data: [4]int{0: 7, 3: 9}}
__p(fmt.Sprint(b.data[0]))
__p(fmt.Sprint(b.data[2]))
__p(fmt.Sprint(b.data[3]))
__check("7\n0\n9")
}
