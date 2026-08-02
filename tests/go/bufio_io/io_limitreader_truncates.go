// vybe-test: go/bufio_io/io_limitreader_truncates
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

func main() { lr := io.LimitReader(strings.NewReader("longtext"), 4)
data, _ := io.ReadAll(lr)
__check(fmt.Sprint(string(data)), "long") }
