// vybe-test: go/sync_package/sync_map_range_accumulates_values
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
m.Store("a", 10)
m.Store("b", 20)
sum := 0
m.Range(func(k, v interface{}) bool { sum += v.(int); return true })
__check(fmt.Sprint(sum), "30") }
