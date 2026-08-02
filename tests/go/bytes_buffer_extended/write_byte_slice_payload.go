// vybe-test: go/bytes_buffer_extended/write_byte_slice_payload
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
b.Write([]byte("xy"))
__check(fmt.Sprint(b.String()), "xy") }
