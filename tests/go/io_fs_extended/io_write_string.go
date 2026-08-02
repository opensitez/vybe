// vybe-test: go/io_fs_extended/io_write_string
// origin: languages/go/tests/go/test_io_fs_extended.rs

package main
import "fmt"
import "bytes"
import "io"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var buf bytes.Buffer
n, _ := io.WriteString(&buf, "go")
__check(fmt.Sprint(n) + " " + fmt.Sprint(buf.String()), "2 go") }
