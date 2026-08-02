// vybe-test: go/io_pipe_copy_tee/copy_chained_multi_reader_source
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
src := io.MultiReader(strings.NewReader("ab"), strings.NewReader("cd"))
_, _ = io.Copy(&dst, src)
__check(fmt.Sprint(dst.String()), "abcd") }
