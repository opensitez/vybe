// vybe-test: go/sync_map_once_extended/sync_map_swap_compile
// origin: languages/go/tests/go/test_sync_map_once_extended.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { var m sync.Map
_, _ = m.Swap("k", 1) }
