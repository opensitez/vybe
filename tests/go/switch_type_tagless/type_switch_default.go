// vybe-test: go/switch_type_tagless/type_switch_default
// origin: languages/go/tests/go/test_switch_type_tagless.rs

package main
import "fmt"
func describe(v interface{}) { switch v.(type) { case int: __p(fmt.Sprint("int")) default: __p(fmt.Sprint("default")) } }
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

func main() { describe(1.5) 
__check("default")
}
