// vybe-test: go/atomic_sync_extended/store_load_int64_defer_after_goroutine
// origin: languages/go/tests/go/test_atomic_sync_extended.rs
// vybe-test-mode: compile

package main
import "sync/atomic"
func main() { var n int64
go func() { atomic.StoreInt64(&n, 5) }()
_ = atomic.LoadInt64(&n) }
