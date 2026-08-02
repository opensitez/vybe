// vybe-test: go/bytes_buffer_extended/readfrom_appends_reader_payload
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
_, _ = b.ReadFrom(strings.NewReader("vybe"))
__check(fmt.Sprint(b.String()), "vybe") }
