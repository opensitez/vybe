// vybe-test: go/bufio_io/reader_read_slice_until_comma
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

func main() { r := bufio.NewReader(strings.NewReader("a,b"))
part, _ := r.ReadSlice(',')
__check(fmt.Sprint(string(part)), "a,") }
