// vybe-test: go/panic_recover_rules/recover_from_panic_in_if_init
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { __p(fmt.Sprint(recover())) }()
if x := 1; x > 0 { panic("if") } }
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

func main() { run() 
__check("if")
}
