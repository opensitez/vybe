// vybe-test: go/io_pipe_copy_tee/copy_n_partial_when_source_shorter
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
n, err := io.CopyN(&dst, strings.NewReader("ab"), 5)
__check(fmt.Sprint(n) + " " + fmt.Sprint(dst.String()) + " " + fmt.Sprint(err != nil), "2 ab true") }
