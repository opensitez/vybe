// vybe-test: go/bytes_buffer_extended/reset_clears_then_rebuilds
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

func main() { var b bytes.Buffer
b.WriteString("old")
b.Reset()
b.WriteString("new")
__check(fmt.Sprint(b.String()), "new") }
