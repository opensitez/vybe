// vybe-test: go/sync_package/pool_put_then_get_returns_stored_value
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

func main() { var p sync.Pool
p.New = func() interface{} { return 1 }
p.Put(9)
__check(fmt.Sprint(p.Get().(int)), "9") }
