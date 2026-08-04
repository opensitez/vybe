// vybe-test: go/variadic_spread/mixed_literals_plus_spread_string_slice
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func join3(a string, b string, rest ...string) int { return len(rest) + len(a) + len(b) }
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

func main() { tail := []string{"c", "d"}
__p(fmt.Sprint(join3("x", "y", tail...)))
__check("4")
}
