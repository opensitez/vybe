// vybe-test: go/sync_map_once_extended/sync_pool_put_interface_values_compile
// origin: languages/go/tests/go/test_sync_map_once_extended.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { var p sync.Pool
p.Put(1)
p.Put("x")
_ = p.Get() }
