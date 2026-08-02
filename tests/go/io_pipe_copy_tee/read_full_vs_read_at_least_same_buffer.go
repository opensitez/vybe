// vybe-test: go/io_pipe_copy_tee/read_full_vs_read_at_least_same_buffer
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

func main() { buf1 := make([]byte, 2)
buf2 := make([]byte, 2)
_, e1 := io.ReadFull(strings.NewReader("xy"), buf1)
_, e2 := io.ReadAtLeast(strings.NewReader("xy"), buf2, 2)
__check(fmt.Sprint(string(buf1)) + " " + fmt.Sprint(string(buf2)) + " " + fmt.Sprint(e1 == nil) + " " + fmt.Sprint(e2 == nil), "xy xy true true") }
