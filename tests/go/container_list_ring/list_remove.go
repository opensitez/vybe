// vybe-test: go/container_list_ring/list_remove
// origin: languages/go/tests/go/test_container_list_ring.rs

package main
import "fmt"
import "container/list"
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

func main() { l := list.New()
e := l.PushBack(9)
l.Remove(e)
__p(fmt.Sprint(l.Len())) 
__check("0")
}
