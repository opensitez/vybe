// vybe-test: go/sync_package/sync_map_delete_removes_key
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

func main() { var m sync.Map
m.Store("d", 1)
m.Delete("d")
_, ok := m.Load("d")
__p(fmt.Sprint(ok)) 
__check("false")
}
