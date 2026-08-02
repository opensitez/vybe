// vybe-test: go/sync_map_once_extended/sync_cond_new_with_locker_compile
// origin: languages/go/tests/go/test_sync_map_once_extended.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { var rw sync.RWMutex
_ = sync.NewCond(&rw) }
