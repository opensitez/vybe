// vybe-test: go/io_pipe_copy_tee/copy_repeated_from_limited_reader
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

func main() { var dst bytes.Buffer
src := io.LimitReader(strings.NewReader("abcdef"), 3)
n, _ := io.Copy(&dst, src)
__check(fmt.Sprint(n) + " " + fmt.Sprint(dst.String()), "3 abc") }
