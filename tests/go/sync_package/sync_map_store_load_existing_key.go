// vybe-test: go/sync_package/sync_map_store_load_existing_key
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
m.Store("k", 42)
v, ok := m.Load("k")
__check(fmt.Sprint(v.(int)), "42")
__check(fmt.Sprint(ok), "true") }
