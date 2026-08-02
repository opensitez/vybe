// vybe-test: go/atomic_sync_extended/swap_int64_concurrent_with_load
// origin: languages/go/tests/go/test_atomic_sync_extended.rs
// vybe-test-mode: compile

package main
import "sync/atomic"
func main() { var n int64
atomic.StoreInt64(&n, 1)
go func() { atomic.SwapInt64(&n, 2) }()
_ = atomic.LoadInt64(&n) }
