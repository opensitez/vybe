// vybe-test: go/container_heap_list/ring_do_single_element
// origin: languages/go/tests/go/test_container_heap_list.rs

package main
import "fmt"
import "container/ring"
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

func main() { r := ring.New(1)
r.Value = 42
count := 0
r.Do(func(v interface{}) { count++; __p(fmt.Sprint(v)) })
__p(fmt.Sprint(count)) 
__check("42\n1")
}
