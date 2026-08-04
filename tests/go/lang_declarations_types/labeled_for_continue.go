// vybe-test: go/lang_declarations_types/labeled_for_continue
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

func main() { L: for i := 0; i < 3; i++ { if i == 1 { continue L }
__p(fmt.Sprint(i)) } 
__check("0\n2")
}
