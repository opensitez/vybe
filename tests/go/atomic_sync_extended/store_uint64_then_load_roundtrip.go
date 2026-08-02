// vybe-test: go/atomic_sync_extended/store_uint64_then_load_roundtrip
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
atomic.StoreUint64(&n, 10000000000)
__check(fmt.Sprint(atomic.LoadUint64(&n)), "10000000000") }
