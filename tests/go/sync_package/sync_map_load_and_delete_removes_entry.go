// vybe-test: go/sync_package/sync_map_load_and_delete_removes_entry
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
m.Store("x", 5)
v, ok := m.LoadAndDelete("x")
_, still := m.Load("x")
__check(fmt.Sprint(v.(int)), "5")
__check(fmt.Sprint(ok), "true")
__check(fmt.Sprint(still), "false") }
