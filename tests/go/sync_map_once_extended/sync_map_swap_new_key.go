// vybe-test: go/sync_map_once_extended/sync_map_swap_new_key
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
prev, loaded := m.Swap("n", 7)
__check(fmt.Sprint(prev == nil), "true")
__check(fmt.Sprint(loaded), "false")
v, _ := m.Load("n")
__check(fmt.Sprint(v.(int)), "7") }
