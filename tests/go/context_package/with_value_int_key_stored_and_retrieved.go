// vybe-test: go/context_package/with_value_int_key_stored_and_retrieved
// origin: languages/go/tests/go/test_context_package.rs

package main
import "fmt"
import "context"
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { type key int
const idKey key = 1
ctx := context.WithValue(context.Background(), idKey, 99)
__p(fmt.Sprint(ctx.Value(idKey).(int))) 
__check("99")
}
