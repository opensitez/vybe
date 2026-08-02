// vybe-test: go/sync_map_once_extended/sync_map_swap_replaces_value
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
m.Store("k", 1)
prev, loaded := m.Swap("k", 2)
__check(fmt.Sprint(prev.(int)), "1")
__check(fmt.Sprint(loaded), "true")
v, _ := m.Load("k")
__check(fmt.Sprint(v.(int)), "2") }
