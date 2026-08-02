// vybe-test: go/io_fs_extended/io_read_byte
// origin: languages/go/tests/go/test_io_fs_extended.rs

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

func main() { r := strings.NewReader("A")
b, err := r.ReadByte()
__check(fmt.Sprint(string(b)) + " " + fmt.Sprint(err == nil), "A true") }
