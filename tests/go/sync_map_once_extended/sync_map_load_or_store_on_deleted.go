// vybe-test: go/sync_map_once_extended/sync_map_load_or_store_on_deleted
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
m.Delete("k")
actual, loaded := m.LoadOrStore("k", 2)
__check(fmt.Sprint(actual.(int)), "2")
__check(fmt.Sprint(loaded), "false") }
