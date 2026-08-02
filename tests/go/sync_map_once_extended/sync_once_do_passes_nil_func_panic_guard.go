// vybe-test: go/sync_map_once_extended/sync_once_do_passes_nil_func_panic_guard
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

func main() { var once sync.Once
ran := false
once.Do(func() { ran = true })
__check(fmt.Sprint(ran), "true") }
