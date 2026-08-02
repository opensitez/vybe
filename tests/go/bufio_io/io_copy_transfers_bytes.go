// vybe-test: go/bufio_io/io_copy_transfers_bytes
// origin: languages/go/tests/go/test_bufio_io.rs

package main
import "fmt"
import "io"
import "bytes"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var dst bytes.Buffer
_, _ = io.Copy(&dst, strings.NewReader("copy"))
__check(fmt.Sprint(dst.String()), "copy") }
