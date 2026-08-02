// vybe-test: go/atomic_sync_extended/atomic_value_store_struct_pointer
// origin: languages/go/tests/go/test_atomic_sync_extended.rs
// vybe-test-mode: compile

package main
import "sync/atomic"
type cfg struct { Port int }
func main() { var v atomic.Value
v.Store(&cfg{Port: 8080})
_ = v.Load().(*cfg) }
