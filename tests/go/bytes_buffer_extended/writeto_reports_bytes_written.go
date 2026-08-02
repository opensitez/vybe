// vybe-test: go/bytes_buffer_extended/writeto_reports_bytes_written
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

func main() { var src bytes.Buffer
src.WriteString("data")
var dst bytes.Buffer
n, _ := src.WriteTo(&dst)
__check(fmt.Sprint(n), "4") }
