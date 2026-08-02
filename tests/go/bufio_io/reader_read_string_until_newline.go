// vybe-test: go/bufio_io/reader_read_string_until_newline
// origin: languages/go/tests/go/test_bufio_io.rs

package main
import "fmt"
import "bufio"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { r := bufio.NewReader(strings.NewReader("hi\n"))
s, _ := r.ReadString('\n')
__check(fmt.Sprint(s), "hi\n") }
