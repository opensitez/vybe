// vybe-test: go/sync_map_once_extended/sync_pool_struct_new_factory
// origin: languages/go/tests/go/test_sync_map_once_extended.rs

package main
import "fmt"
import "sync"
type buf struct { data []byte }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var p sync.Pool
p.New = func() interface{} { return &buf{data: make([]byte, 0, 8)} }
b := p.Get().(*buf)
__check(fmt.Sprint(cap(b.data)), "8") }
