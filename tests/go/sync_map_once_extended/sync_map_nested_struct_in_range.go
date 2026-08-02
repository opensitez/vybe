// vybe-test: go/sync_map_once_extended/sync_map_nested_struct_in_range
// origin: languages/go/tests/go/test_sync_map_once_extended.rs

package main
import "fmt"
import "sync"
type pair struct { a int
b int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var m sync.Map
m.Store(1, pair{a: 2, b: 3})
sum := 0
m.Range(func(k, v interface{}) bool { p := v.(pair); sum = p.a + p.b; return true })
__check(fmt.Sprint(sum), "5") }
