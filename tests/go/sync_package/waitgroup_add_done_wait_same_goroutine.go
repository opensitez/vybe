// vybe-test: go/sync_package/waitgroup_add_done_wait_same_goroutine
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

func main() { var wg sync.WaitGroup
wg.Add(1)
wg.Done()
wg.Wait()
__check(fmt.Sprint("done"), "done") }
