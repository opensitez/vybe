// vybe-test: go/io_pipe_copy_tee/tee_reader_partial_read_duplicates_prefix
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs

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

func main() { var side bytes.Buffer
tr := io.TeeReader(strings.NewReader("abcd"), &side)
buf := make([]byte, 2)
tr.Read(buf)
__check(fmt.Sprint(string(buf)) + " " + fmt.Sprint(side.String()), "ab ab") }
