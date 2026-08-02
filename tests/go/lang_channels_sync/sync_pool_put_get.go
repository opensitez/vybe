// vybe-test: go/lang_channels_sync/sync_pool_put_get
// origin: languages/go/tests/go/test_lang_channels_sync.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { p := sync.Pool{New: func() interface{} { return 0 }}
p.Put(1)
_ = p.Get() }
