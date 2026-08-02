// vybe-test: go/context_cancel_deadline/context_with_timeout_deadline_exceeded_after_sleep
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

func main() { ctx, cancel := context.WithTimeout(context.Background(), time.Millisecond)
defer cancel()
time.Sleep(2 * time.Millisecond)
__check(fmt.Sprint(ctx.Err() == context.DeadlineExceeded), "true") }
