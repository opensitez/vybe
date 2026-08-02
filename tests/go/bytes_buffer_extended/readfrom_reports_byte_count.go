// vybe-test: go/bytes_buffer_extended/readfrom_reports_byte_count
// origin: languages/go/tests/go/test_bytes_buffer_extended.rs

package main
import "fmt"
import "bytes"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var b bytes.Buffer
n, _ := b.ReadFrom(strings.NewReader("four"))
__check(fmt.Sprint(n), "4") }
