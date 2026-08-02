// vybe-test: go/sync_package/pool_new_invoked_on_empty_get
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
p.New = func() interface{} { return 7 }
__check(fmt.Sprint(p.Get().(int)), "7") }
