// vybe-test: go/atomic_sync_extended/typed_uint64_add_and_load
// origin: languages/go/tests/go/test_atomic_sync_extended.rs
// vybe-test-mode: compile

package main
import "sync/atomic"
func main() { var v atomic.Uint64
v.Add(10)
_ = v.Load() }
