// vybe-test: go/sync_package/waitgroup_add_multiple_before_single_done
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
wg.Add(3)
wg.Done()
wg.Done()
wg.Done()
wg.Wait()
__check(fmt.Sprint(0), "0") }
