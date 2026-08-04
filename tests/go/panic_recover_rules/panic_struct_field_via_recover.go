// vybe-test: go/panic_recover_rules/panic_struct_field_via_recover
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
type err struct { code int }
func run() { defer func() { e := recover().(err)
__p(fmt.Sprint(e.code)) }()
panic(err{code: 5}) }
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
__check("5")
}
