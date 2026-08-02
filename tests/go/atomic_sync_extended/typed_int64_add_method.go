// vybe-test: go/atomic_sync_extended/typed_int64_add_method
// origin: languages/go/tests/go/test_atomic_sync_extended.rs
// vybe-test-mode: compile

package main
import "sync/atomic"
func main() { var v atomic.Int64
v.Add(3)
_ = v.Load() }
