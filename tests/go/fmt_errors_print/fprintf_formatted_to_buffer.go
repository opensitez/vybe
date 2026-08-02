// vybe-test: go/fmt_errors_print/fprintf_formatted_to_buffer
// origin: languages/go/tests/go/test_fmt_errors_print.rs

package main
import "fmt"
import "bytes"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var buf bytes.Buffer
fmt.Fprintf(&buf, "id=%d", 7)
__check(fmt.Sprint(buf.String()), "id=7") }
