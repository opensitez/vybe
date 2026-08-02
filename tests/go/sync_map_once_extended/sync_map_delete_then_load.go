// vybe-test: go/sync_map_once_extended/sync_map_delete_then_load
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
m.Store("d", 1)
m.Delete("d")
_, ok := m.Load("d")
__check(fmt.Sprint(ok), "false") }
