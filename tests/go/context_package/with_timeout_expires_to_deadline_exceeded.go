// vybe-test: go/context_package/with_timeout_expires_to_deadline_exceeded
// origin: languages/go/tests/go/test_context_package.rs

package main
import "fmt"
import "context"
import "time"
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

func main() { ctx, cancel := context.WithTimeout(context.Background(), time.Millisecond)
defer cancel()
time.Sleep(2 * time.Millisecond)
__p(fmt.Sprint(ctx.Err() == context.DeadlineExceeded)) 
__check("true")
}
