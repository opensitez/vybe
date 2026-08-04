// vybe-test: go/type_switch_extended/type_switch_single_int_case_only
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case int: __p(fmt.Sprint("int")) } }
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

func main() { tag(7) 
__check("int")
}
