// vybe-test: go/context_package/with_value_missing_key_returns_nil
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

func main() { ctx := context.WithValue(context.Background(), "a", 1)
__check(fmt.Sprint(ctx.Value("b") == nil), "true") }
