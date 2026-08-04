// vybe-test: go/atomic_sync_extended/swap_int64_leaves_new_value_in_memory
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
atomic.StoreInt64(&n, 3)
atomic.SwapInt64(&n, 9)
__p(fmt.Sprint(atomic.LoadInt64(&n))) 
__check("9")
}
