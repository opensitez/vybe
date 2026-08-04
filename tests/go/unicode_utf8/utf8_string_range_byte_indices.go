// vybe-test: go/unicode_utf8/utf8_string_range_byte_indices
// origin: languages/go/tests/go/test_unicode_utf8.rs

package main
import "fmt"
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

func main() { first, second := -1, -1
step := 0
for i, _ := range "a世" { if step == 0 { first = i }
if step == 1 { second = i }
step++ }
__p(fmt.Sprint(first))
__p(fmt.Sprint(second)) 
__check("0\n1")
}
