// vybe-test: go/sync_package/pool_concurrent_get_put_cycle
// origin: languages/go/tests/go/test_sync_package.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { var p sync.Pool
p.New = func() interface{} { return 0 }
go func() { p.Put(p.Get()) }()
_ = p.Get() }
