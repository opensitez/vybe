// vybe-test: go/sync_map_once_extended/sync_map_range_collect_keys
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
m.Store("x", 1)
m.Store("y", 2)
keys := 0
m.Range(func(k, v interface{}) bool { keys++; return true })
__check(fmt.Sprint(keys), "2") }
