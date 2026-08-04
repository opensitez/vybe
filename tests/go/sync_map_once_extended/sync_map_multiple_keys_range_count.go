// vybe-test: go/sync_map_once_extended/sync_map_multiple_keys_range_count
// origin: languages/go/tests/go/test_sync_map_once_extended.rs

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
m.Store("a", 1)
m.Store("b", 2)
m.Store("c", 3)
n := 0
m.Range(func(k, v interface{}) bool { n++; return true })
__p(fmt.Sprint(n)) 
__check("3")
}
