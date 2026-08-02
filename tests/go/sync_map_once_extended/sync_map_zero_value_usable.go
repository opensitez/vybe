// vybe-test: go/sync_map_once_extended/sync_map_zero_value_usable
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
m.Store(0, 0)
v, ok := m.Load(0)
__check(fmt.Sprint(v.(int)), "0")
__check(fmt.Sprint(ok), "true") }
