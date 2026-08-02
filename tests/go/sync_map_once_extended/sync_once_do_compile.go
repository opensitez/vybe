// vybe-test: go/sync_map_once_extended/sync_once_do_compile
// origin: languages/go/tests/go/test_sync_map_once_extended.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { var once sync.Once
once.Do(func() {}) }
