// vybe-test: go/context_cancel_deadline/context_with_value_missing_key_nil
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

func main() { ctx := context.WithValue(context.Background(), "x", 1)
__check(fmt.Sprint(ctx.Value("y") == nil), "true") }
