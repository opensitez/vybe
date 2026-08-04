// vybe-test: go/struct_embedding_extra/anonymous_struct_in_slice_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

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

func main() { values := []struct { name string }{{name: "vybe"}}
__p(fmt.Sprint(values[0].name))
__check("vybe")
}
