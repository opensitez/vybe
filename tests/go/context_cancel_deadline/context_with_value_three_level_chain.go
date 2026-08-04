// vybe-test: go/context_cancel_deadline/context_with_value_three_level_chain
// origin: languages/go/tests/go/test_context_cancel_deadline.rs

package main
import "fmt"
import "context"
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

func main() { c1 := context.WithValue(context.Background(), "a", 1)
c2 := context.WithValue(c1, "b", 2)
c3 := context.WithValue(c2, "c", 3)
__p(fmt.Sprint(c3.Value("a").(int)))
__p(fmt.Sprint(c3.Value("b").(int)))
__p(fmt.Sprint(c3.Value("c").(int))) 
__check("1\n2\n3")
}
