// vybe-test: go/bufio_scanner_extended/reader_peek_after_read_string
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs
// vybe-test-mode: compile

package main
import "bufio"
import "strings"
func main() { r := bufio.NewReader(strings.NewReader("a\nb"))
_, _ = r.ReadString('\n')
_, _ = r.Peek(1) }
