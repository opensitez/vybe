// vybe-test: go/bytes_buffer_extended/read_reduces_unread_len
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
b.WriteString("abcd")
buf := make([]byte, 2)
_, _ = b.Read(buf)
__check(fmt.Sprint(b.Len()), "2") }
