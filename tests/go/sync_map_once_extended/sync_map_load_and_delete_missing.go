// vybe-test: go/sync_map_once_extended/sync_map_load_and_delete_missing
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
_, ok := m.LoadAndDelete("z")
__check(fmt.Sprint(ok), "false") }
