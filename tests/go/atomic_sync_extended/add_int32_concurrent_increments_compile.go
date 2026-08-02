// vybe-test: go/atomic_sync_extended/add_int32_concurrent_increments_compile
// origin: languages/go/tests/go/test_atomic_sync_extended.rs
// vybe-test-mode: compile

package main
import "sync/atomic"
func main() { var n int32
go func() { atomic.AddInt32(&n, 1) }()
go func() { atomic.AddInt32(&n, 2) }()
_ = atomic.LoadInt32(&n) }
