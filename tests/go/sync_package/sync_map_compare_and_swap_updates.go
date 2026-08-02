// vybe-test: go/sync_package/sync_map_compare_and_swap_updates
// origin: languages/go/tests/go/test_sync_package.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { var m sync.Map
m.Store("k", 1)
_ = m.CompareAndSwap("k", 1, 2) }
