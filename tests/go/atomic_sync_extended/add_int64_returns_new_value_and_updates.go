// vybe-test: go/atomic_sync_extended/add_int64_returns_new_value_and_updates
// origin: languages/go/tests/go/test_atomic_sync_extended.rs

package main
import "fmt"
import "sync/atomic"
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

func main() { var n int64
atomic.StoreInt64(&n, 10)
__p(fmt.Sprint(atomic.AddInt64(&n, 5)))
__p(fmt.Sprint(atomic.LoadInt64(&n))) 
__check("15\n15")
}
