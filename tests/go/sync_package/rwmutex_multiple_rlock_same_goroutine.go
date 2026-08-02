// vybe-test: go/sync_package/rwmutex_multiple_rlock_same_goroutine
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
rw.RLock()
rw.RLock()
__check(fmt.Sprint("ok"), "ok")
rw.RUnlock()
rw.RUnlock() }
