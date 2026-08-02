// vybe-test: go/bufio_io/io_nopcloser_preserves_read
// origin: languages/go/tests/go/test_bufio_io.rs

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

func main() { rc := io.NopCloser(strings.NewReader("wrap"))
data, _ := io.ReadAll(rc)
__check(fmt.Sprint(string(data)), "wrap") }
