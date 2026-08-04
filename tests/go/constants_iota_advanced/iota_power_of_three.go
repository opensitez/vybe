// vybe-test: go/constants_iota_advanced/iota_power_of_three
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( P0 = 1; P1 = 3 * iota; P2 )
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

func main() { __p(fmt.Sprint(P0))
__p(fmt.Sprint(P1))
__p(fmt.Sprint(P2)) 
__check("1\n3\n6")
}
