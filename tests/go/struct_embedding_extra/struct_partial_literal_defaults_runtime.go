// vybe-test: go/struct_embedding_extra/struct_partial_literal_defaults_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type item struct { count int
name string }
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

func main() { value := item{name: "go"}
__p(fmt.Sprint(value.count))
__p(fmt.Sprint(value.name))
__check("0\ngo")
}
