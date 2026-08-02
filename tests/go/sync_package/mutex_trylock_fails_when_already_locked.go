// vybe-test: go/sync_package/mutex_trylock_fails_when_already_locked
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

func main() { var mu sync.Mutex
mu.Lock()
__check(fmt.Sprint(mu.TryLock()), "false")
mu.Unlock() }
