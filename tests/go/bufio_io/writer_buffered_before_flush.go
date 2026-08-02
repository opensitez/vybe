// vybe-test: go/bufio_io/writer_buffered_before_flush
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
w.WriteString("ab")
__check(fmt.Sprint(w.Buffered()), "2")
w.Flush()
__check(fmt.Sprint(buf.String()), "ab") }
