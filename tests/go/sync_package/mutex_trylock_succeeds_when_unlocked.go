// vybe-test: go/sync_package/mutex_trylock_succeeds_when_unlocked
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
__check(fmt.Sprint(mu.TryLock()), "true")
mu.Unlock() }
