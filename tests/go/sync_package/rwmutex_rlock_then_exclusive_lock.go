// vybe-test: go/sync_package/rwmutex_rlock_then_exclusive_lock
// origin: languages/go/tests/go/test_sync_package.rs

package main
import "fmt"
import "sync"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var rw sync.RWMutex
v := 0
rw.RLock()
v = 1
rw.RUnlock()
rw.Lock()
v = 2
rw.Unlock()
__check(fmt.Sprint(v), "2") }
