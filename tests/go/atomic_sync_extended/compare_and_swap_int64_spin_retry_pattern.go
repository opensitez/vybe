// vybe-test: go/atomic_sync_extended/compare_and_swap_int64_spin_retry_pattern
// origin: languages/go/tests/go/test_atomic_sync_extended.rs
// vybe-test-mode: compile

package main
import "sync/atomic"
func main() { var n int64
atomic.StoreInt64(&n, 0)
for !atomic.CompareAndSwapInt64(&n, 0, 1) { }
_ = n }
