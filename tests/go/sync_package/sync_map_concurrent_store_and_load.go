// vybe-test: go/sync_package/sync_map_concurrent_store_and_load
// origin: languages/go/tests/go/test_sync_package.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { var m sync.Map
go func() { m.Store("k", 1) }()
_, _ = m.Load("k") }
