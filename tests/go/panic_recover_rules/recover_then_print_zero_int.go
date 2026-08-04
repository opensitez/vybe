// vybe-test: go/panic_recover_rules/recover_then_print_zero_int
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { recover()
var n int
__p(fmt.Sprint(n)) }()
panic("stop") }
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
__check("0")
}
