// vybe-test: go/sync_map_once_extended/sync_map_in_struct_field_compile
// origin: languages/go/tests/go/test_sync_map_once_extended.rs
// vybe-test-mode: compile

package main
import "sync"
type cache struct { m sync.Map }
func main() { var c cache
c.m.Store(1, 2) }
