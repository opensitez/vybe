// vybe-test: go/lang_channels_sync/sync_rwlock
// origin: languages/go/tests/go/test_lang_channels_sync.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { var rw sync.RWMutex
rw.RLock()
rw.RUnlock() }
