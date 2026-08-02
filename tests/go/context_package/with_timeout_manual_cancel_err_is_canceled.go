// vybe-test: go/context_package/with_timeout_manual_cancel_err_is_canceled
// origin: languages/go/tests/go/test_context_package.rs

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

func main() { ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
cancel()
__check(fmt.Sprint(ctx.Err() == context.Canceled), "true") }
