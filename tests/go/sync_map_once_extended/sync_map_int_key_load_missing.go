// vybe-test: go/sync_map_once_extended/sync_map_int_key_load_missing
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
_, ok := m.Load(99)
__check(fmt.Sprint(ok), "false") }
