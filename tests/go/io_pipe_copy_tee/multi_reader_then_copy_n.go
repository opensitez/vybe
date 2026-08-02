// vybe-test: go/io_pipe_copy_tee/multi_reader_then_copy_n
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

func main() { mr := io.MultiReader(strings.NewReader("ab"), strings.NewReader("cd"))
var dst bytes.Buffer
n, _ := io.CopyN(&dst, mr, 3)
__check(fmt.Sprint(n) + " " + fmt.Sprint(dst.String()), "3 abc") }
