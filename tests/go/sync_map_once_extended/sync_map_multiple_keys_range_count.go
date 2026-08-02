// vybe-test: go/sync_map_once_extended/sync_map_multiple_keys_range_count
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

func main() { var m sync.Map
m.Store("a", 1)
m.Store("b", 2)
m.Store("c", 3)
n := 0
m.Range(func(k, v interface{}) bool { n++; return true })
__check(fmt.Sprint(n), "3") }
