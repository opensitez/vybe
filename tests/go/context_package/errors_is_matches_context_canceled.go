// vybe-test: go/context_package/errors_is_matches_context_canceled
// origin: languages/go/tests/go/test_context_package.rs

package main
import "fmt"
import "context"
import "errors"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ctx, cancel := context.WithCancel(context.Background())
cancel()
__check(fmt.Sprint(errors.Is(ctx.Err(), context.Canceled)), "true") }
