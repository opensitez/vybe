// vybe-test: go/sync_package/sync_map_load_or_store_keeps_existing
// origin: languages/go/tests/go/test_sync_package.rs

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
actual, loaded := m.LoadOrStore("a", 99)
__check(fmt.Sprint(actual.(int)), "1")
__check(fmt.Sprint(loaded), "true") }
