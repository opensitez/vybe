// vybe-test: go/sync_package/rwmutex_trylock_after_write_lock
// origin: languages/go/tests/go/test_sync_package.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { var rw sync.RWMutex
rw.Lock()
_ = rw.TryRLock()
rw.Unlock() }
