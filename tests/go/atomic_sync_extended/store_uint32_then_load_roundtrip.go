// vybe-test: go/atomic_sync_extended/store_uint32_then_load_roundtrip
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

func main() { var n uint32
atomic.StoreUint32(&n, 65535)
__check(fmt.Sprint(atomic.LoadUint32(&n)), "65535") }
