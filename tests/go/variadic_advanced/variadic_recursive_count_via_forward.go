// vybe-test: go/variadic_advanced/variadic_recursive_count_via_forward
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func depth(level int, tags ...string) int { if level == 0 { return len(tags) }
return depth(level-1, tags...) }
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

func main() { __p(fmt.Sprint(depth(2, "a", "b", "c"))) 
__check("3")
}
