// vybe-test: go/lang_channels_sync/sync_map_store_load
// origin: languages/go/tests/go/test_lang_channels_sync.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { var m sync.Map
m.Store("k", 1)
_, _ = m.Load("k") }
