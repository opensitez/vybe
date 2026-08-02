// vybe-test: go/sync_map_once_extended/sync_cond_signal_compile
// origin: languages/go/tests/go/test_sync_map_once_extended.rs
// vybe-test-mode: compile

package main
import "sync"
func main() { var mu sync.Mutex
cond := sync.NewCond(&mu)
cond.Signal() }
