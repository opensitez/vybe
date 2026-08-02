// vybe-test: go/bufio_io/writer_reset_rebinds_output
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

func main() { var a bytes.Buffer
var b bytes.Buffer
w := bufio.NewWriter(&a)
w.WriteString("old")
w.Reset(&b)
w.WriteString("new")
w.Flush()
__check(fmt.Sprint(a.String()), "")
__check(fmt.Sprint(b.String()), "new") }
