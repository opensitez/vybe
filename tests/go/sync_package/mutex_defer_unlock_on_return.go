// vybe-test: go/sync_package/mutex_defer_unlock_on_return
// origin: languages/go/tests/go/test_sync_package.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { var mu sync.Mutex
mu.Lock()
defer mu.Unlock()
_ = 1 }
