// vybe-test: go/lang_channels_sync/sync_mutex_lock_unlock
// origin: languages/go/tests/go/test_lang_channels_sync.rs

package main
import "fmt"
import "sync"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var m sync.Mutex
m.Lock()
m.Unlock()
__check(fmt.Sprint("ok"), "ok") }
