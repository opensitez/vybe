// vybe-test: go/sync_map_once_extended/sync_map_compare_and_delete_wrong_old
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
m.Store("k", 1)
deleted := m.CompareAndDelete("k", 2)
_, ok := m.Load("k")
__check(fmt.Sprint(deleted), "false")
__check(fmt.Sprint(ok), "true") }
