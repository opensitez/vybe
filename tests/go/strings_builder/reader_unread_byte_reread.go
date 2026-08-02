// vybe-test: go/strings_builder/reader_unread_byte_reread
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
b1, _ := r.ReadByte()
_ = b1
r.UnreadByte()
b2, _ := r.ReadByte()
__check(fmt.Sprint(string(b2)), "g") }
