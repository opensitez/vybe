// vybe-test: go/sync_package/sync_map_load_missing_key
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
_, ok := m.Load("missing")
__check(fmt.Sprint(ok), "false") }
