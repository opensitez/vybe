// vybe-test: go/sync_package/mutex_trylock_fails_when_already_locked
// origin: languages/go/tests/go/test_sync_package.rs

package main
import "fmt"
import "sync"
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { var mu sync.Mutex
mu.Lock()
__p(fmt.Sprint(mu.TryLock()))
mu.Unlock() 
__check("false")
}
