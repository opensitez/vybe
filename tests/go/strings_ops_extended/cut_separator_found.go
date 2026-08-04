// vybe-test: go/strings_ops_extended/cut_separator_found
// origin: languages/go/tests/go/test_strings_ops_extended.rs

package main
import "fmt"
import "strings"
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

func main() { before, after, found := strings.Cut("hello,world", ",")
__p(fmt.Sprint(before))
__p(fmt.Sprint(after))
__p(fmt.Sprint(found)) 
__check("hello\nworld\ntrue")
}
