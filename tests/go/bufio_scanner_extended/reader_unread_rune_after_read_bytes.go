// vybe-test: go/bufio_scanner_extended/reader_unread_rune_after_read_bytes
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs
// vybe-test-mode: compile

package main
import "bufio"
import "strings"
func main() { r := bufio.NewReader(strings.NewReader("ab"))
_, _ = r.ReadBytes('a')
_ = r.UnreadRune() }
