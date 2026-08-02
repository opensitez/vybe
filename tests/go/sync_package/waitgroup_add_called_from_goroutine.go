// vybe-test: go/sync_package/waitgroup_add_called_from_goroutine
// origin: languages/go/tests/go/test_sync_package.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { var wg sync.WaitGroup
ch := make(chan struct{})
go func() { wg.Add(1)
close(ch) }()
<-ch
go func() { wg.Done() }()
wg.Wait() }
