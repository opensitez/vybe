// vybe-test: go/bufio_io/reader_peek_without_advancing
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

func main() { r := bufio.NewReader(strings.NewReader("abc"))
peek, _ := r.Peek(2)
__check(fmt.Sprint(string(peek)), "ab")
b, _ := r.ReadByte()
__check(fmt.Sprint(string(b)), "a") }
