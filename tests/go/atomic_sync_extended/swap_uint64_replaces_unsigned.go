// vybe-test: go/atomic_sync_extended/swap_uint64_replaces_unsigned
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

func main() { var n uint64
atomic.StoreUint64(&n, 5)
__check(fmt.Sprint(atomic.SwapUint64(&n, 8)), "5")
__check(fmt.Sprint(atomic.LoadUint64(&n)), "8") }
