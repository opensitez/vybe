// vybe-test: go/sync_map_once_extended/sync_map_range_callback_compile
// origin: languages/go/tests/go/test_sync_map_once_extended.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { var m sync.Map
m.Range(func(k, v interface{}) bool { return true }) }
