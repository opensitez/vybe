// vybe-test: go/sync_package/mutex_serial_increment_under_lock
// origin: languages/go/tests/go/test_sync_package.rs

package main
import "fmt"
import "sync"
func main() { var mu sync.Mutex
n := 0
for i := 0; i < 5; i++ { mu.Lock()
n++
mu.Unlock() }
fmt.Println(n) }
