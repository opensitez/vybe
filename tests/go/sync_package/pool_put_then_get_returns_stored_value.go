// vybe-test: go/sync_package/pool_put_then_get_returns_stored_value
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

func main() { var p sync.Pool
p.New = func() interface{} { return 1 }
p.Put(9)
__p(fmt.Sprint(p.Get().(int))) 
__check("9")
}
