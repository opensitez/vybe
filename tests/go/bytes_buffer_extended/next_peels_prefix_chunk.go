// vybe-test: go/bytes_buffer_extended/next_peels_prefix_chunk
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
b.WriteString("abcdef")
chunk := b.Next(2)
__check(fmt.Sprint(string(chunk)), "ab")
__check(fmt.Sprint(b.String()), "cdef") }
