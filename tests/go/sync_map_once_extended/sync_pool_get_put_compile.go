// vybe-test: go/sync_map_once_extended/sync_pool_get_put_compile
// origin: languages/go/tests/go/test_sync_map_once_extended.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { var p sync.Pool
p.New = func() interface{} { return 0 }
p.Put(p.Get()) }
