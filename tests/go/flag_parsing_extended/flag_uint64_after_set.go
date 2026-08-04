// vybe-test: go/flag_parsing_extended/flag_uint64_after_set
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

func main() { big := flag.Uint64("big", 0, "")
_ = flag.Set("big", "5000")
__p(fmt.Sprint(*big)) 
__check("5000")
}
