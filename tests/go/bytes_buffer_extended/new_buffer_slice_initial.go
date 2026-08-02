// vybe-test: go/bytes_buffer_extended/new_buffer_slice_initial
// origin: languages/go/tests/go/test_bytes_buffer_extended.rs

package main
import "fmt"
import "bytes"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { b := bytes.NewBuffer([]byte("init"))
__check(fmt.Sprint(b.String()), "init") }
