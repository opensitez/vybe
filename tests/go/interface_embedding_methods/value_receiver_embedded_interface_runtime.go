// vybe-test: go/interface_embedding_methods/value_receiver_embedded_interface_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type scaler interface { scale() int }
type resizable interface { scaler }
type icon struct { n int }
func (i icon) scale() int { return i.n * 2 }
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

func main() { var r resizable = icon{n: 6}
__p(fmt.Sprint(r.scale())) 
__check("12")
}
