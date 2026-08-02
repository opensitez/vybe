// vybe-test: go/context_package/with_value_int_key_stored_and_retrieved
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

func main() { type key int
const idKey key = 1
ctx := context.WithValue(context.Background(), idKey, 99)
__check(fmt.Sprint(ctx.Value(idKey).(int)), "99") }
