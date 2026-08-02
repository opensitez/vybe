// vybe-test: go/sync_package/mutex_goroutines_increment_shared_counter
// origin: languages/go/tests/go/test_sync_package.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { var mu sync.Mutex
n := 0
for i := 0; i < 3; i++ { go func() { mu.Lock()
n++
mu.Unlock() }() }
mu.Lock()
mu.Unlock() }
