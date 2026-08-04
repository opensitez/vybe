// vybe-test: go/flag_parsing_extended/flag_new_flagset_visit_all_after_define
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

func main() { fs := flag.NewFlagSet("fs", flag.ContinueOnError)
_ = fs.Bool("on", false, "")
n := 0
fs.VisitAll(func(f *flag.Flag) { n++ })
__p(fmt.Sprint(n)) 
__check("1")
}
