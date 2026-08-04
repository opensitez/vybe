// vybe-test: go/variadic_advanced/variadic_string_max_length
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func longest(words ...string) int { m := 0
for _, w := range words { if len(w) > m { m = len(w) } }
return m }
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

func main() { __p(fmt.Sprint(longest("go", "vybe", "a"))) 
__check("4")
}
