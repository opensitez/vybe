// vybe-test: go/lang_channels_sync/sync_once_do
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

func main() { var o sync.Once
n := 0
o.Do(func() { n++ })
o.Do(func() { n++ })
__check(fmt.Sprint(n), "1") }
