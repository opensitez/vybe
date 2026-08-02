// vybe-test: go/atomic_sync_extended/add_uint64_concurrent_counter
// origin: languages/go/tests/go/test_atomic_sync_extended.rs
// vybe-test-mode: compile

package main
import "sync/atomic"
func main() { var n uint64
for i := 0; i < 3; i++ { go func() { atomic.AddUint64(&n, 1) }() }
_ = atomic.LoadUint64(&n) }
