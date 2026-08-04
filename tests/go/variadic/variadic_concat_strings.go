// vybe-test: go/variadic/variadic_concat_strings
// origin: languages/go/tests/go/test_variadic.rs

package main
import "fmt"
func joinAll(sep string, parts ...string) string { r := ""
i := 0
for _, p := range parts { if i > 0 { r = r + sep }
r = r + p
i++ }
return r } var __buf string

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

func main() { __p(fmt.Sprint(joinAll("-", "a", "b", "c")))
__check("a-b-c")
}
