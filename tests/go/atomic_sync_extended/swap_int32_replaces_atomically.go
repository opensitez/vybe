// vybe-test: go/atomic_sync_extended/swap_int32_replaces_atomically
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

func main() { var n int32
atomic.StoreInt32(&n, 11)
__check(fmt.Sprint(atomic.SwapInt32(&n, 22)), "11")
__check(fmt.Sprint(atomic.LoadInt32(&n)), "22") }
