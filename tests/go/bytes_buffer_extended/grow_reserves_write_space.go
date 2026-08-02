// vybe-test: go/bytes_buffer_extended/grow_reserves_write_space
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
b.Grow(64)
b.WriteString("x")
__check(fmt.Sprint(b.Len()), "1") }
