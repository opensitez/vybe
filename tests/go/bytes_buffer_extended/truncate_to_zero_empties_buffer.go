// vybe-test: go/bytes_buffer_extended/truncate_to_zero_empties_buffer
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
b.WriteString("x")
b.Truncate(0)
__check(fmt.Sprint(b.Len()), "0") }
