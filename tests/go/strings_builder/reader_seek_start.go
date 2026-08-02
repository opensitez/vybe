// vybe-test: go/strings_builder/reader_seek_start
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
_, _ = r.ReadByte()
pos, _ := r.Seek(0, 0)
__check(fmt.Sprint(pos), "0") }
