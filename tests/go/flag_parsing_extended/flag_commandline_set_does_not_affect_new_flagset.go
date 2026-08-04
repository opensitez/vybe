// vybe-test: go/flag_parsing_extended/flag_commandline_set_does_not_affect_new_flagset
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

func main() { _ = flag.String("shared", "cmd", "")
fs := flag.NewFlagSet("other", flag.ContinueOnError)
local := fs.String("shared", "local", "")
__p(fmt.Sprint(*local)) 
__check("local")
}
