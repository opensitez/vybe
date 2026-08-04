// vybe-test: go/atomic_sync_extended/swap_uint64_replaces_unsigned
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

func main() { var n uint64
atomic.StoreUint64(&n, 5)
__p(fmt.Sprint(atomic.SwapUint64(&n, 8)))
__p(fmt.Sprint(atomic.LoadUint64(&n))) 
__check("5\n8")
}
