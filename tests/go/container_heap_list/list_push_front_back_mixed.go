// vybe-test: go/container_heap_list/list_push_front_back_mixed
// origin: languages/go/tests/go/test_container_heap_list.rs

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
l.PushBack(2)
l.PushFront(1)
l.PushBack(3)
__p(fmt.Sprint(l.Front().Value))
__p(fmt.Sprint(l.Back().Value))
__p(fmt.Sprint(l.Len())) 
__check("1\n3\n3")
}
