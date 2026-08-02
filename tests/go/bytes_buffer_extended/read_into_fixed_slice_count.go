// vybe-test: go/bytes_buffer_extended/read_into_fixed_slice_count
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
n, _ := b.Read(buf)
__check(fmt.Sprint(n), "2")
__check(fmt.Sprint(string(buf)), "ab") }
