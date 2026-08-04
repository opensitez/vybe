// vybe-test: go/function_types_advanced/method_for_each_with_index_callback
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type batch struct { items []int }
func (b batch) forEach(visit func(int, int)) { for i, v := range b.items { visit(i, v) } }
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

func main() { sum := 0
batch{items: []int{2, 3, 4}}.forEach(func(i int, v int) { sum += v })
__p(fmt.Sprint(sum)) 
__check("9")
}
