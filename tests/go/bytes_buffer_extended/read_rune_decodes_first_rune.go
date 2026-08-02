// vybe-test: go/bytes_buffer_extended/read_rune_decodes_first_rune
// origin: languages/go/tests/go/test_bytes_buffer_extended.rs

package main
import "fmt"
import "bytes"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var b bytes.Buffer
b.WriteString("日lang")
r, _, _ := b.ReadRune()
__check(fmt.Sprint(string(r)), "日") }
