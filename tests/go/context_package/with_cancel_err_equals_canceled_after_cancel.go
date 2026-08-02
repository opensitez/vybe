// vybe-test: go/context_package/with_cancel_err_equals_canceled_after_cancel
// origin: languages/go/tests/go/test_context_package.rs

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
