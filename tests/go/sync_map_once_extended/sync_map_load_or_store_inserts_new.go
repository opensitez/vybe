// vybe-test: go/sync_map_once_extended/sync_map_load_or_store_inserts_new
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
actual, loaded := m.LoadOrStore("b", 3)
__check(fmt.Sprint(actual.(int)), "3")
__check(fmt.Sprint(loaded), "false") }
