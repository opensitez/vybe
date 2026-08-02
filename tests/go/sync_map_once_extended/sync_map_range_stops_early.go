// vybe-test: go/sync_map_once_extended/sync_map_range_stops_early
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
m.Store(1, 1)
m.Store(2, 2)
count := 0
m.Range(func(k, v interface{}) bool { count++; return false })
__check(fmt.Sprint(count), "1") }
