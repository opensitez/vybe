// vybe-test: go/sync_map_once_extended/sync_once_nested_do_only_outer_counts
// origin: languages/go/tests/go/test_sync_map_once_extended.rs

package main
import "fmt"
import "sync"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var outer sync.Once
var inner sync.Once
n := 0
outer.Do(func() { inner.Do(func() { n++ }); inner.Do(func() { n++ }) })
__check(fmt.Sprint(n), "1") }
