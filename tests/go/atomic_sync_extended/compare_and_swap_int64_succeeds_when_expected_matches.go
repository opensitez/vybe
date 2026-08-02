// vybe-test: go/atomic_sync_extended/compare_and_swap_int64_succeeds_when_expected_matches
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
atomic.StoreInt64(&n, 1)
__check(fmt.Sprint(atomic.CompareAndSwapInt64(&n, 1, 2)), "true")
__check(fmt.Sprint(atomic.LoadInt64(&n)), "2") }
