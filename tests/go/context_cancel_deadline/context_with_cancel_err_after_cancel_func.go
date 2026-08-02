// vybe-test: go/context_cancel_deadline/context_with_cancel_err_after_cancel_func
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

func main() { ctx, cancel := context.WithCancel(context.Background())
cancel()
__check(fmt.Sprint(ctx.Err() == context.Canceled), "true") }
