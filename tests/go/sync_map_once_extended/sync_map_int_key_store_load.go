// vybe-test: go/sync_map_once_extended/sync_map_int_key_store_load
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
m.Store(1, "one")
v, ok := m.Load(1)
__check(fmt.Sprint(v.(string)), "one")
__check(fmt.Sprint(ok), "true") }
