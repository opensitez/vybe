// vybe-test: go/sync_map_once_extended/sync_map_compare_and_swap_success
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
swapped := m.CompareAndSwap("k", 1, 2)
v, _ := m.Load("k")
__check(fmt.Sprint(swapped), "true")
__check(fmt.Sprint(v.(int)), "2") }
