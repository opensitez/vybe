// vybe-test: go/sync_package/waitgroup_waits_for_spawned_goroutines
// origin: languages/go/tests/go/test_sync_package.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { var wg sync.WaitGroup
wg.Add(1)
go func() { defer wg.Done()
_ = 1 }()
wg.Wait() }
