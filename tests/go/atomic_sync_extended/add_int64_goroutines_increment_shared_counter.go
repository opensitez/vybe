// vybe-test: go/atomic_sync_extended/add_int64_goroutines_increment_shared_counter
// origin: languages/go/tests/go/test_atomic_sync_extended.rs
// vybe-test-mode: compile

package main
import "sync/atomic"
func main() { var n int64
for i := 0; i < 5; i++ { go func() { atomic.AddInt64(&n, 1) }() }
_ = atomic.LoadInt64(&n) }
