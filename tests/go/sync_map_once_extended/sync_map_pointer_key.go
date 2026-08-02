// vybe-test: go/sync_map_once_extended/sync_map_pointer_key
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

func main() { key := new(int)
*key = 42
var m sync.Map
m.Store(key, "ptr")
v, ok := m.Load(key)
__check(fmt.Sprint(v.(string)), "ptr")
__check(fmt.Sprint(ok), "true") }
