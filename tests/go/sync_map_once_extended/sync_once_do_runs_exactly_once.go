// vybe-test: go/sync_map_once_extended/sync_once_do_runs_exactly_once
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
n := 0
f := func() { n++ }
once.Do(f)
once.Do(f)
__check(fmt.Sprint(n), "1") }
