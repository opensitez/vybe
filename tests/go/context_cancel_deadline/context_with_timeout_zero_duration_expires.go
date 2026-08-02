// vybe-test: go/context_cancel_deadline/context_with_timeout_zero_duration_expires
// origin: languages/go/tests/go/test_context_cancel_deadline.rs

package main
import "fmt"
import "context"
import "time"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ctx, cancel := context.WithTimeout(context.Background(), 0)
defer cancel()
__check(fmt.Sprint(ctx.Err() == context.DeadlineExceeded), "true") }
