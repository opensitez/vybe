// vybe-test: go/bufio_scanner_extended/reader_peek_with_small_reader_buffer
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs
// vybe-test-mode: compile

package main
import "bufio"
import "strings"
func main() { r := bufio.NewReaderSize(strings.NewReader("abcd"), 2)
_, _ = r.Peek(3) }
