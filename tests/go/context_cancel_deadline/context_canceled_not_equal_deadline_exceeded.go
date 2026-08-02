// vybe-test: go/context_cancel_deadline/context_canceled_not_equal_deadline_exceeded
// origin: languages/go/tests/go/test_context_cancel_deadline.rs

package main
import "fmt"
import "context"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(context.Canceled != context.DeadlineExceeded), "true") }
