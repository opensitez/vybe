// vybe-test: go/strings_builder/reader_read_byte_first
// origin: languages/go/tests/go/test_strings_builder.rs

package main
import "fmt"
import "strings"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { r := strings.NewReader("go")
b, _ := r.ReadByte()
__check(fmt.Sprint(string(b)), "g") }
