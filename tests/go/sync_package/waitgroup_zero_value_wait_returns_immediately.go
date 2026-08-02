// vybe-test: go/sync_package/waitgroup_zero_value_wait_returns_immediately
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
wg.Wait()
__check(fmt.Sprint("ready"), "ready") }
