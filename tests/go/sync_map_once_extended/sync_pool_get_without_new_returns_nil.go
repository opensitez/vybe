// vybe-test: go/sync_map_once_extended/sync_pool_get_without_new_returns_nil
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

func main() { var p sync.Pool
__check(fmt.Sprint(p.Get() == nil), "true") }
