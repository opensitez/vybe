// vybe-test: go/sync_map_once_extended/sync_map_load_and_delete_then_load_or_store
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
m.LoadAndDelete("k")
actual, loaded := m.LoadOrStore("k", 3)
__check(fmt.Sprint(actual.(int)), "3")
__check(fmt.Sprint(loaded), "false") }
