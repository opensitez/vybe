// vybe-test: go/io_pipe_copy_tee/multi_reader_single_byte_reads
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

func main() { mr := io.MultiReader(strings.NewReader("12"), strings.NewReader("34"))
buf := make([]byte, 1)
mr.Read(buf)
__check(fmt.Sprint(string(buf)), "1")
mr.Read(buf)
__check(fmt.Sprint(string(buf)), "2") }
