// vybe-test: go/io_pipe_copy_tee/multi_reader_three_segments
// origin: languages/go/tests/go/test_io_pipe_copy_tee.rs

package main
import "fmt"
import "io"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { mr := io.MultiReader(strings.NewReader("a"), strings.NewReader("b"), strings.NewReader("c"))
data, _ := io.ReadAll(mr)
__check(fmt.Sprint(string(data)), "abc") }
