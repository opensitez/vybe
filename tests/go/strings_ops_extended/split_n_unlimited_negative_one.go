// vybe-test: go/strings_ops_extended/split_n_unlimited_negative_one
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

func main() { parts := strings.SplitN("one:two:three", ":", -1)
__p(fmt.Sprint(len(parts))) 
__check("3")
}
