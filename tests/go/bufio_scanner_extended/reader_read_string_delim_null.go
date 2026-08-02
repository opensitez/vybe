// vybe-test: go/bufio_scanner_extended/reader_read_string_delim_null
// origin: languages/go/tests/go/test_bufio_scanner_extended.rs
// vybe-test-mode: compile

package main
import "bufio"
import "strings"
func main() { r := bufio.NewReader(strings.NewReader("a\x00b"))
_, _ = r.ReadString('\x00') }
