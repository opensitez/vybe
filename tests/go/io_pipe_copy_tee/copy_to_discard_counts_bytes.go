// vybe-test: go/io_pipe_copy_tee/copy_to_discard_counts_bytes
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

func main() { n, _ := io.Copy(io.Discard, strings.NewReader("discard"))
__check(fmt.Sprint(n), "7") }
