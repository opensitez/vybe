// vybe-test: go/sync_map_once_extended/sync_map_range_sums_int_values
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
m.Store(1, 10)
m.Store(2, 20)
sum := 0
m.Range(func(k, v interface{}) bool { sum += v.(int); return true })
__check(fmt.Sprint(sum), "30") }
