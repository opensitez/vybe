// vybe-test: go/panic_recover_rules/recover_after_defer_modifies_named_return
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() (n int) { defer func() { recover()
n = 9 }()
panic("p")
return 1 }
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

func main() { __p(fmt.Sprint(run())) 
__check("9")
}
