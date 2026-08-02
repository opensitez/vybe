// vybe-test: go/atomic_sync_extended/atomic_value_store_load_int
// origin: languages/go/tests/go/test_atomic_sync_extended.rs
// vybe-test-mode: compile

package main
import "sync/atomic"
func main() { var v atomic.Value
v.Store(42)
_ = v.Load().(int) }
