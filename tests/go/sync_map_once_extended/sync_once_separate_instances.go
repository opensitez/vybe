// vybe-test: go/sync_map_once_extended/sync_once_separate_instances
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

func main() { var a sync.Once
var b sync.Once
n := 0
a.Do(func() { n++ })
b.Do(func() { n++ })
__check(fmt.Sprint(n), "2") }
