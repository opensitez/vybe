// vybe-test: go/bytes_buffer_extended/read_byte_consumes_first
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
b.WriteString("go")
ch, _ := b.ReadByte()
__check(fmt.Sprint(string(ch)), "g") }
