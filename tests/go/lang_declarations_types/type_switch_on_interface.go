// vybe-test: go/lang_declarations_types/type_switch_on_interface
// origin: languages/go/tests/go/test_lang_declarations_types.rs

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

func main() { switch any("x").(type) { case string: __p(fmt.Sprint("s"))
default: __p(fmt.Sprint("d")) } 
__check("s")
}
