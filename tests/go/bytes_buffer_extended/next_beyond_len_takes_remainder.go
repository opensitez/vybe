// vybe-test: go/bytes_buffer_extended/next_beyond_len_takes_remainder
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
b.WriteString("hi")
chunk := b.Next(10)
__check(fmt.Sprint(string(chunk)), "hi")
__check(fmt.Sprint(b.Len()), "0") }
