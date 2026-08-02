// vybe-test: go/sync_map_once_extended/sync_once_do_with_closure_capture
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
sum := 0
once.Do(func() { sum = 10 })
once.Do(func() { sum = 99 })
__check(fmt.Sprint(sum), "10") }
