// vybe-test: go/defer_panic_variants/recover_captures_struct_field_via_type_assertion
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
type stop struct { code int }
func run() { defer func() { value := recover()
if err, ok := value.(stop); ok { __p(fmt.Sprint(err.code)) } }()
panic(stop{code: 42}) }
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
__check("42")
}
