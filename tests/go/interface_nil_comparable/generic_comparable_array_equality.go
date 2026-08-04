// vybe-test: go/interface_nil_comparable/generic_comparable_array_equality
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func equalArray(left [2]int, right [2]int) bool { return left == right }
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

func main() { __p(fmt.Sprint(equalArray([2]int{1, 2}, [2]int{1, 2})))
__p(fmt.Sprint(equalArray([2]int{1, 2}, [2]int{2, 1}))) 
__check("true\nfalse")
}
