// vybe-test: go/sync_map_once_extended/sync_pool_reuse_after_put
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
p.New = func() interface{} { return 1 }
first := p.Get().(int)
p.Put(first + 10)
second := p.Get().(int)
__check(fmt.Sprint(second), "11") }
