// vybe-test: go/sync_package/sync_map_swap_replaces_value
// origin: languages/go/tests/go/test_sync_package.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { var m sync.Map
m.Store("k", 1)
_, _ = m.Swap("k", 2) }
