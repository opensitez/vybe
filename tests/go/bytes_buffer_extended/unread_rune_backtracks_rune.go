// vybe-test: go/bytes_buffer_extended/unread_rune_backtracks_rune
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
b.WriteString("日x")
b.ReadRune()
b.UnreadRune()
r, _, _ := b.ReadRune()
__check(fmt.Sprint(string(r)), "日") }
