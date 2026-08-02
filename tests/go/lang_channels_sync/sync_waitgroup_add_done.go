// vybe-test: go/lang_channels_sync/sync_waitgroup_add_done
// origin: languages/go/tests/go/test_lang_channels_sync.rs

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
__check(fmt.Sprint(1), "1") }
