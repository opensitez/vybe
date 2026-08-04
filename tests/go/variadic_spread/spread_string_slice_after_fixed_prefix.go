// vybe-test: go/variadic_spread/spread_string_slice_after_fixed_prefix
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func tag(prefix string, words ...string) { for _, w := range words { __p(fmt.Sprint(prefix + w)) } }
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

func main() { rest := []string{"go", "vybe"}
tag(">", rest...)
__check(">go\n>vybe")
}
