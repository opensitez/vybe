// vybe-test: go/type_switch_extended/type_switch_return_from_case_via_helper
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func size(v interface{}) int { switch v.(type) { case string: return len(v.(string))
case int: return v.(int)
default: return 0 } }
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

func main() { __p(fmt.Sprint(size("abc")))
__p(fmt.Sprint(size(10))) 
__check("3\n10")
}
