// vybe-test: go/atomic_sync_extended/add_int64_sequential_increments
// origin: languages/go/tests/go/test_atomic_sync_extended.rs

package main
import "fmt"
import "sync/atomic"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var n int64
atomic.AddInt64(&n, 1)
atomic.AddInt64(&n, 2)
atomic.AddInt64(&n, 3)
__check(fmt.Sprint(atomic.LoadInt64(&n)), "6") }
