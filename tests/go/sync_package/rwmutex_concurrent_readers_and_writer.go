// vybe-test: go/sync_package/rwmutex_concurrent_readers_and_writer
// origin: languages/go/tests/go/test_sync_package.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { var rw sync.RWMutex
rw.RLock()
go func() { rw.RLock()
rw.RUnlock() }()
rw.RUnlock() }
