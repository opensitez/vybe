// vybe-test: go/embedding_promotion_extended/embed_pointer_vs_value_mutation_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

package main
import "fmt"
type cell struct { n int }
type wrapValue struct { cell }
type wrapPtr struct { *cell }
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

func main() { v := wrapValue{cell: cell{n: 1}}
p := wrapPtr{cell: &cell{n: 1}}
v.cell.n = 9
p.n = 8
__p(fmt.Sprint(v.n))
__p(fmt.Sprint(p.n)) 
__check("9\n8")
}
