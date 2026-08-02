// vybe-test: go/bufio_io/writer_write_string_after_flush
// origin: languages/go/tests/go/test_bufio_io.rs

package main
import "fmt"
import "bufio"
import "bytes"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var buf bytes.Buffer
w := bufio.NewWriter(&buf)
w.WriteString("vybe")
w.Flush()
__check(fmt.Sprint(buf.String()), "vybe") }
