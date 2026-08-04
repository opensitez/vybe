// vybe-test: go/fmt_errors_print/sscanf_hex_integer
// origin: languages/go/tests/go/test_fmt_errors_print.rs

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

func main() { var n int
c, _ := fmt.Sscanf("ff", "%x", &n)
__p(fmt.Sprint(c) + " " + fmt.Sprint(n)) 
__check("1 255")
}
