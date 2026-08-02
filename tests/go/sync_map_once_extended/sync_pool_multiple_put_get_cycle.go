// vybe-test: go/sync_map_once_extended/sync_pool_multiple_put_get_cycle
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
p.Put(1)
p.Put(2)
a := p.Get().(int)
b := p.Get().(int)
__check(fmt.Sprint(a + b), "3") }
