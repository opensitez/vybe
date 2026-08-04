// vybe-test: go/context_cancel_deadline/context_with_value_missing_key_nil
// origin: languages/go/tests/go/test_context_cancel_deadline.rs

package main
import "fmt"
import "context"
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

func main() { ctx := context.WithValue(context.Background(), "x", 1)
__p(fmt.Sprint(ctx.Value("y") == nil)) 
__check("true")
}
