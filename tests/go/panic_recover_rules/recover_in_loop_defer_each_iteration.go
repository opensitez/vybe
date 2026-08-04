// vybe-test: go/panic_recover_rules/recover_in_loop_defer_each_iteration
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { for i := 0; i < 2; i++ { defer func(n int) { if recover() != nil { __p(fmt.Sprint(n)) } }(i) }
panic("loop") }
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

func main() { defer func() { recover() }()
run() 
__check("1")
}
