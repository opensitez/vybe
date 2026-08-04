// vybe-test: go/sync_package/rwmutex_rlock_then_exclusive_lock
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

func main() { var rw sync.RWMutex
v := 0
rw.RLock()
v = 1
rw.RUnlock()
rw.Lock()
v = 2
rw.Unlock()
__p(fmt.Sprint(v)) 
__check("2")
}
