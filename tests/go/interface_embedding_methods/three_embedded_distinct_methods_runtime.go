// vybe-test: go/interface_embedding_methods/three_embedded_distinct_methods_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type alpha interface { a() int }
type beta interface { b() int }
type gamma interface { c() int }
type combo interface { alpha
beta
gamma }
type triple struct{}
func (triple) a() int { return 1 }
func (triple) b() int { return 2 }
func (triple) c() int { return 3 }
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

func main() { var value combo = triple{}
__p(fmt.Sprint(value.a()))
__p(fmt.Sprint(value.b()))
__p(fmt.Sprint(value.c())) 
__check("1\n2\n3")
}
