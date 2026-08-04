// vybe-test: go/blank_identifier_extended/blank_multi_assign_three_returns
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
func triple() (int, int, int) { return 1, 2, 3 }
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

func main() { _, mid, _ := triple()
__p(fmt.Sprint(mid)) 
__check("2")
}
