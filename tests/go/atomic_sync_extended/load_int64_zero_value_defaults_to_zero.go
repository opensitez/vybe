// vybe-test: go/atomic_sync_extended/load_int64_zero_value_defaults_to_zero
// origin: languages/go/tests/go/test_atomic_sync_extended.rs

package main
import "fmt"
import "sync/atomic"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var n int64
__check(fmt.Sprint(atomic.LoadInt64(&n)), "0") }
