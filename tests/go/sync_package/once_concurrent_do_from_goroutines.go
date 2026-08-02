// vybe-test: go/sync_package/once_concurrent_do_from_goroutines
// origin: languages/go/tests/go/test_sync_package.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { var once sync.Once
for i := 0; i < 3; i++ { go once.Do(func() {}) }
once.Do(func() {}) }
