// vybe-test: go/bufio_io/io_readall_entire_content
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

func main() { data, _ := io.ReadAll(strings.NewReader("full"))
__check(fmt.Sprint(string(data)), "full") }
