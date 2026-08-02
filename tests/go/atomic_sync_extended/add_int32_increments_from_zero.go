// vybe-test: go/atomic_sync_extended/add_int32_increments_from_zero
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
__check(fmt.Sprint(atomic.AddInt32(&n, 3)), "3")
__check(fmt.Sprint(atomic.LoadInt32(&n)), "3") }
