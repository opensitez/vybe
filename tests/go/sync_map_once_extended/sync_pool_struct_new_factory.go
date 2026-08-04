// vybe-test: go/sync_map_once_extended/sync_pool_struct_new_factory
// origin: languages/go/tests/go/test_sync_map_once_extended.rs

package main
import "fmt"
import "sync"
type buf struct { data []byte }
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
p.New = func() interface{} { return &buf{data: make([]byte, 0, 8)} }
b := p.Get().(*buf)
__p(fmt.Sprint(cap(b.data))) 
__check("8")
}
