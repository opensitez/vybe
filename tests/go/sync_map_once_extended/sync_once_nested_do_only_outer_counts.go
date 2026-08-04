// vybe-test: go/sync_map_once_extended/sync_once_nested_do_only_outer_counts
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

func main() { var outer sync.Once
var inner sync.Once
n := 0
outer.Do(func() { inner.Do(func() { n++ }); inner.Do(func() { n++ }) })
__p(fmt.Sprint(n)) 
__check("1")
}
