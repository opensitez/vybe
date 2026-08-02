// vybe-test: go/context_package/with_value_returns_stored_string_value
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

func main() { ctx := context.WithValue(context.Background(), "token", "abc")
__check(fmt.Sprint(ctx.Value("token").(string)), "abc") }
