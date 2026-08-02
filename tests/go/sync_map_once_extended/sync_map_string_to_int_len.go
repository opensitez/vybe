// vybe-test: go/sync_map_once_extended/sync_map_string_to_int_len
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
m.Store("hello", len("hello"))
v, _ := m.Load("hello")
__check(fmt.Sprint(v.(int)), "5") }
