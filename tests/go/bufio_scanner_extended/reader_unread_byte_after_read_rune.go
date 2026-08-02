// vybe-test: go/bufio_scanner_extended/reader_unread_byte_after_read_rune
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs
// vybe-test-mode: compile

package main
import "bufio"
import "strings"
func main() { r := bufio.NewReader(strings.NewReader("go"))
_, _, _ = r.ReadRune()
_ = r.UnreadByte() }
