// vybe-test: go/embedding_promotion_extended/two_embedded_same_field_name_requires_qualifier_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type left struct { id int }
type right struct { id int }
type pair struct { left
right }
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

func main() { p := pair{left: left{id: 1}, right: right{id: 2}}
__p(fmt.Sprint(p.left.id))
__p(fmt.Sprint(p.right.id)) 
__check("1\n2")
}
