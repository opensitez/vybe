// vybe-test: go/constants_iota_advanced/iota_shift_left_per_step
// origin: languages/go/tests/go/test_constants_iota_advanced.rs

package main
import "fmt"
const ( Bit0 = 1 << iota; Bit1; Bit2; Bit3 )
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

func main() { __p(fmt.Sprint(Bit0))
__p(fmt.Sprint(Bit3)) 
__check("1\n8")
}
