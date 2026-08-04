// vybe-test: go/constants_iota_advanced/iota_or_chain_three_flags
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( F1 = 1 << iota; F2; F3 )
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

func main() { mask := F1 | F2
__p(fmt.Sprint(mask))
__p(fmt.Sprint(mask | F3)) 
__check("3\n7")
}
