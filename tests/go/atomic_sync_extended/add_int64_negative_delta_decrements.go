// vybe-test: go/atomic_sync_extended/add_int64_negative_delta_decrements
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
atomic.StoreInt64(&n, 20)
__check(fmt.Sprint(atomic.AddInt64(&n, -8)), "12")
__check(fmt.Sprint(atomic.LoadInt64(&n)), "12") }
