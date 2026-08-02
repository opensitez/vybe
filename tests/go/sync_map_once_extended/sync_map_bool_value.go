// vybe-test: go/sync_map_once_extended/sync_map_bool_value
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
m.Store(true, "yes")
v, ok := m.Load(true)
__check(fmt.Sprint(v.(string)), "yes")
__check(fmt.Sprint(ok), "true") }
