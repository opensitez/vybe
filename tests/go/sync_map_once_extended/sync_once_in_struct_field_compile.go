// vybe-test: go/sync_map_once_extended/sync_once_in_struct_field_compile
// origin: languages/go/tests/go/test_sync_map_once_extended.rs
// vybe-test-mode: compile

package main
import "sync"
type holder struct { once sync.Once }
func main() { var h holder
h.once.Do(func() {}) }
