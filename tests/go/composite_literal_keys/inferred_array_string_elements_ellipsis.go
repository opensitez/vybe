// vybe-test: go/composite_literal_keys/inferred_array_string_elements_ellipsis
// origin: languages/go/tests/go/test_composite_literal_keys.rs

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

func main() { words := [...]string{"go", "vybe", "keys"}
__p(fmt.Sprint(len(words)))
__p(fmt.Sprint(words[2]))
__check("3\nkeys")
}
