// vybe-test: go/struct_embedding_advanced/value_embed_pointer_receiver_promotion_runtime
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs

package main
import "fmt"
type inner struct { n int }
func (i *inner) double() { i.n = i.n * 2 }
type outer struct { inner }
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

func main() { value := outer{inner: inner{n: 4}}
value.double()
__p(fmt.Sprint(value.n))
__check("8")
}
