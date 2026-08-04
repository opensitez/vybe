// vybe-test: go/flag_parsing_extended/flag_set_uint_max
// origin: languages/go/tests/go/test_flag_parsing_extended.rs

package main
import "fmt"
import "flag"
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

func main() { v := flag.Uint("v", 0, "")
_ = flag.Set("v", "4294967295")
__p(fmt.Sprint(*v)) 
__check("4294967295")
}
