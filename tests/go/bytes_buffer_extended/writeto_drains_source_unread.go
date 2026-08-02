// vybe-test: go/bytes_buffer_extended/writeto_drains_source_unread
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

func main() { var src bytes.Buffer
src.WriteString("go")
var dst bytes.Buffer
_, _ = src.WriteTo(&dst)
__check(fmt.Sprint(src.Len()), "0") }
