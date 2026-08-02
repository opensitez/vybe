// vybe-test: go/sync_map_once_extended/sync_map_struct_value_roundtrip
// origin: languages/go/tests/go/test_sync_map_once_extended.rs

package main
import "fmt"
import "sync"
type item struct { n int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var m sync.Map
m.Store("k", item{n: 8})
v, ok := m.Load("k")
__check(fmt.Sprint(v.(item).n), "8")
__check(fmt.Sprint(ok), "true") }
